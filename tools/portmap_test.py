"""
portmap_test.py - port mapping (libplum, cluster J part 2) against a SIMULATED ROUTER on 127.0.0.1 only.

  python -I tools\\portmap_test.py [--exe PATH] [scenario ...]     (pcp natpmp upnp fail timeout off; default all)

Nothing here may reach the real router or the LAN's UPnP devices (the user's rule):
* Warzone.exe runs with the TEST-ONLY override WZ_TEST_PORTMAPPING_GATEWAY=127.0.0.1 (lib/netplay's libplum port,
  plum_net_c: the default gateway is that address, so PCP / NAT-PMP go to 127.0.0.1:5351) and
  WZ_TEST_PORTMAPPING_SSDP_PORT (plum_upnp_c's SSDP search goes unicast to 127.0.0.1:<port>, never to the multicast
  group; Windows' SSDP service holds 127.0.0.1:1900, so the fake answers on another port). The override takes only a
  loopback address. The run is refused unless both are set; the config's portMapping=1 is written only then, and
  its lobby is http://127.0.0.1:9 (nothing listens: no game is listed on the public lobby).
* The fake router binds 127.0.0.1 only: a PCP server or a NAT-PMP one on UDP 5351, an SSDP responder, an IGD's
  description and its SOAP control (GetExternalIPAddress, AddPortMapping, DeletePortMapping) over HTTP.
* After each run the log must show the override at every gateway lookup and the fake's address as the gateway.

Each scenario: the exe (default build\\j_PM\\bin\\Warzone.exe, the fork's private build) goes Multi Player > Host Game
(NETinitPortMapping: the mapping is created) > Start Hosting (NEThostGame waits for the mapping), waits, then
Return (the hosting cancelled) and Return (the screen left: NETremRedirects removes the mapping), waits, quits.
* pcp: the fake is a PCP server: the mapping succeeds with its external address; leaving sends a MAP of lifetime 0.
* natpmp: the fake answers PCP with UNSUPP_VERSION version 0 and is a NAT-PMP server: the mapping succeeds.
  libplum's natpmp_impl_check_epoch_time compares the new epoch with itself (not the previous one), so any
  response 3 s or more after the last reads as a router reset: leaving resets the protocol and drops the mapping
  without an unmapping request (the official build does the same); the test expects that.
* upnp: no PCP server (ICMP port unreachable: "UDP recvfrom failed, errno=10054"); the SSDP search finds the fake
  IGD: the mapping succeeds (AddPortMapping); leaving sends DeletePortMapping.
* fail: nothing answers: the official build's failure lines and console text.
* off: portMapping=0 (ui_test.py's default): no port mapping at all ("disabled by user"), hosting as before.
* timeout: the IGD answers its third SSDP search and never answers AddPortMapping: libplum's mapping takes 10 s
  after a discovery of 0.75 s, so PortMappingManager's 10 s timeout comes first ("(timeout)"); libplum's failure
  then finds no request ("no active request found in PortMappingManager, probably timed out").
Results in build/portmap_test/<scenario>/ (log, screenshots, the fake's requests) and build/portmap_test/results.txt.
"""
import glob
import http.server
import os
import shutil
import socket
import socketserver
import subprocess
import sys
import threading
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
WORK = os.path.join(ROOT, "build", "portmap_test")
DEFAULT_EXE = os.path.join(ROOT, "build", "j_PM", "bin", "Warzone.exe")
EXTERNAL_IP = "203.0.113.7"         # TEST-NET-3 (documentation): the fake router's "public" address
LOOPBACK = "127.0.0.1"
PCP_PORT = 5351
GAME_PORT = 2100                    # main picks a free port from 2150 (the texts below say 2100: P())

UI_SCRIPT = """# portmap_test (fork PM): Multi Player, Host Game (NETinitPortMapping), Start Hosting (NEThostGame); the
# mapping's result; Return (the hosting cancelled), Return (the screen left: NETremRedirects); the unmapping; quit
# (the Return arrow at the top left of the options: Escape only takes the chat box's focus)
wait 90
click 512 400
wait 60
click 512 365
wait 120
shot pm_lobby
click 365 455
wait 960
shot pm_hosted
log leaving
click 245 85
wait 120
shot pm_cancelled
click 245 85
wait 120
shot pm_left
wait 600
log done
wait 5
quit
"""


def free_port(kind):
    s = socket.socket(socket.AF_INET, kind)
    s.bind((LOOPBACK, 0))
    port = s.getsockname()[1]
    s.close()
    return port


class FakeRouter:
    """The simulated router: binds 127.0.0.1 only. mode: pcp, natpmp, upnp, fail, timeout."""

    def __init__(self, mode):
        self.mode = mode
        self.requests = []
        self.stop = False
        self.t0 = time.time()
        self.threads = []
        self.sockets = []
        self.ssdp_port = free_port(socket.SOCK_DGRAM)
        self.http_port = free_port(socket.SOCK_STREAM)
        if mode in ("pcp", "natpmp"):
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.bind((LOOPBACK, PCP_PORT))
            s.settimeout(0.2)
            self.sockets.append(s)
            self.threads.append(threading.Thread(target=self.pcp_loop, args=(s,), daemon=True))
        if mode in ("upnp", "timeout"):
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.bind((LOOPBACK, self.ssdp_port))
            s.settimeout(0.2)
            self.sockets.append(s)
            self.threads.append(threading.Thread(target=self.ssdp_loop, args=(s,), daemon=True))
        if mode in ("upnp", "timeout"):
            router = self

            class Handler(http.server.BaseHTTPRequestHandler):
                protocol_version = "HTTP/1.0"

                def log_message(self, *a):
                    pass

                def do_GET(self):
                    router.record("http GET %s" % self.path)
                    if self.path == "/rootDesc.xml":
                        self.reply(200, router.description())
                    else:
                        self.reply(404, "")

                def do_POST(self):
                    body = self.rfile.read(int(self.headers.get("Content-Length", "0"))).decode("utf-8", "replace")
                    action = self.headers.get("SOAPAction", "")
                    router.record("http POST %s SOAPAction=%s %s" % (self.path, action, body[body.find("<s:Body>"):]))
                    if action.endswith("#GetExternalIPAddress"):
                        self.reply(200, router.soap("GetExternalIPAddressResponse",
                                                    "<NewExternalIPAddress>%s</NewExternalIPAddress>" % EXTERNAL_IP))
                    elif action.endswith("#AddPortMapping") and router.mode == "timeout":
                        router.record("http: AddPortMapping never answered")
                        while not router.stop:
                            time.sleep(0.1)
                    elif action.endswith("#AddPortMapping"):
                        self.reply(200, router.soap("AddPortMappingResponse", ""))
                    elif action.endswith("#DeletePortMapping"):
                        self.reply(200, router.soap("DeletePortMappingResponse", ""))
                    else:
                        self.reply(500, "")

                def reply(self, code, text):
                    data = text.encode("utf-8")
                    self.send_response(code)
                    self.send_header("Content-Type", "text/xml; charset=\"utf-8\"")
                    self.send_header("Content-Length", str(len(data)))
                    self.send_header("Connection", "close")
                    self.end_headers()
                    self.wfile.write(data)

            self.httpd = socketserver.ThreadingTCPServer((LOOPBACK, self.http_port), Handler)
            self.httpd.daemon_threads = True
            self.threads.append(threading.Thread(target=self.httpd.serve_forever, daemon=True))
        for t in self.threads:
            t.start()

    def record(self, text):
        self.requests.append("%7.3f %s" % (time.time() - self.t0, text))

    def epoch(self):
        return int(time.time() - self.t0) + 1000

    # PCP (RFC 6887) and NAT-PMP (RFC 6886) on UDP 5351
    def pcp_loop(self, s):
        while not self.stop:
            try:
                data, addr = s.recvfrom(2048)
            except socket.timeout:
                continue
            except OSError:
                continue
            if self.mode == "pcp":
                self.pcp(s, data, addr)
            else:
                self.natpmp(s, data, addr)

    def pcp(self, s, data, addr):
        if len(data) < 24 or data[0] != 2:
            self.record("pcp: ignored %r" % data[:8])
            return
        opcode = data[1]
        raw = data[4:8]
        le, be = int.from_bytes(raw, "little"), int.from_bytes(raw, "big")
        if opcode == 0:
            self.record("pcp ANNOUNCE from %s:%d" % addr)
            s.sendto(bytes([2, 0x80, 0, 0]) + (0).to_bytes(4, "big") + self.epoch().to_bytes(4, "big") + bytes(12), addr)
        elif opcode == 1 and len(data) >= 60:
            nonce, proto = data[24:36], data[36]
            internal = int.from_bytes(data[40:42], "big")
            sugg_port = int.from_bytes(data[42:44], "big")
            # libplum writes the request's lifetime in host order (little-endian): le is what it meant
            self.record("pcp MAP from %s:%d protocol=%d internal=%d suggested_port=%d lifetime(le)=%d lifetime(be)=%d"
                        % (addr[0], addr[1], proto, internal, sugg_port, le, be))
            lifetime = 0 if le == 0 else 7200
            external = sugg_port or internal
            body = (nonce + bytes([proto, 0, 0, 0]) + internal.to_bytes(2, "big") + external.to_bytes(2, "big") +
                    bytes(10) + b"\xff\xff" + socket.inet_aton(EXTERNAL_IP))
            s.sendto(bytes([2, 0x81, 0, 0]) + lifetime.to_bytes(4, "big") + self.epoch().to_bytes(4, "big") +
                     bytes(12) + body, addr)
        else:
            self.record("pcp: opcode %d ignored" % opcode)

    def natpmp(self, s, data, addr):
        if len(data) >= 2 and data[0] == 2:
            # a PCP request to a NAT-PMP server: UNSUPP_VERSION with version 0
            self.record("natpmp: PCP request (opcode %d): UNSUPP_VERSION" % data[1])
            s.sendto(bytes([0, 0x80 | data[1], 0, 1]) + self.epoch().to_bytes(4, "big") + bytes(4), addr)
            return
        if len(data) < 2 or data[0] != 0:
            return
        op = data[1]
        if op == 0:
            self.record("natpmp ANNOUNCE from %s:%d" % addr)
            s.sendto(bytes([0, 0x80, 0, 0]) + self.epoch().to_bytes(4, "big") + socket.inet_aton(EXTERNAL_IP), addr)
        elif op in (1, 2) and len(data) >= 12:
            internal = int.from_bytes(data[4:6], "big")
            sugg = int.from_bytes(data[6:8], "big")
            lifetime = int.from_bytes(data[8:12], "big")
            self.record("natpmp MAP %s from %s:%d internal=%d suggested_port=%d lifetime=%d"
                        % ("UDP" if op == 1 else "TCP", addr[0], addr[1], internal, sugg, lifetime))
            answer = min(lifetime, 7200)
            s.sendto(bytes([0, 0x80 | op, 0, 0]) + self.epoch().to_bytes(4, "big") + internal.to_bytes(2, "big") +
                     (sugg or internal).to_bytes(2, "big") + answer.to_bytes(4, "big"), addr)

    # SSDP (unicast here) and the IGD
    def ssdp_loop(self, s):
        while not self.stop:
            try:
                data, addr = s.recvfrom(4096)
            except socket.timeout:
                continue
            except OSError:
                continue
            text = data.decode("latin-1")
            self.record("ssdp from %s:%d: %s" % (addr[0], addr[1], text.replace("\r\n", " | ")))
            self.searches = getattr(self, "searches", 0) + 1
            if self.mode == "timeout" and self.searches <= 2:
                continue                # timeout: libplum's discovery takes its third search (0.75 s)
            if "M-SEARCH" in text and "InternetGatewayDevice" in text:
                reply = ("HTTP/1.1 200 OK\r\nCACHE-CONTROL: max-age=120\r\n"
                         "ST: urn:schemas-upnp-org:device:InternetGatewayDevice:1\r\n"
                         "USN: uuid:00000000-0000-0000-0000-00000000f00d::urn:schemas-upnp-org:device:InternetGatewayDevice:1\r\n"
                         "EXT:\r\nSERVER: portmap_test/1.0 UPnP/1.1 fake-igd/1.0\r\n"
                         "LOCATION: http://%s:%d/rootDesc.xml\r\n\r\n" % (LOOPBACK, self.http_port))
                s.sendto(reply.encode("latin-1"), addr)

    def description(self):
        return ('<?xml version="1.0"?>\r\n<root xmlns="urn:schemas-upnp-org:device-1-0"><specVersion><major>1</major>'
                '<minor>0</minor></specVersion><device><deviceType>urn:schemas-upnp-org:device:InternetGatewayDevice:1'
                '</deviceType><friendlyName>portmap_test fake IGD</friendlyName><deviceList><device><deviceType>'
                'urn:schemas-upnp-org:device:WANDevice:1</deviceType><deviceList><device><deviceType>'
                'urn:schemas-upnp-org:device:WANConnectionDevice:1</deviceType><serviceList><service>'
                '<serviceType>urn:schemas-upnp-org:service:WANIPConnection:1</serviceType>'
                '<serviceId>urn:upnp-org:serviceId:WANIPConn1</serviceId><controlURL>/ctl/IPConn</controlURL>'
                '<eventSubURL>/evt/IPConn</eventSubURL><SCPDURL>/WANIPCn.xml</SCPDURL></service></serviceList>'
                '</device></deviceList></device></deviceList></device></root>\r\n')

    @staticmethod
    def soap(name, inner):
        return ('<?xml version="1.0"?>\r\n<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" '
                's:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:%s '
                'xmlns:u="urn:schemas-upnp-org:service:WANIPConnection:1">%s</u:%s></s:Body></s:Envelope>\r\n'
                % (name, inner, name))

    def close(self):
        self.stop = True
        if self.mode in ("upnp", "timeout"):
            self.httpd.shutdown()
            self.httpd.server_close()
        for t in self.threads:
            t.join(2)
        for s in self.sockets:
            s.close()


def run_scenario(mode, exe):
    work = os.path.join(WORK, mode)
    if os.path.exists(work):
        shutil.rmtree(work)
    cfg = os.path.join(work, "config")
    os.makedirs(cfg)
    script = os.path.join(work, "portmap.txt")
    open(script, "w").write(UI_SCRIPT)
    router = FakeRouter(mode)
    env = os.environ.copy()
    env["WZ_TEST_PORTMAPPING_GATEWAY"] = LOOPBACK
    env["WZ_TEST_PORTMAPPING_SSDP_PORT"] = str(router.ssdp_port)
    # the safety rule: no run with port mapping on unless the override sends everything to the fake on 127.0.0.1
    assert env.get("WZ_TEST_PORTMAPPING_GATEWAY") == LOOPBACK and env.get("WZ_TEST_PORTMAPPING_SSDP_PORT")
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        f.write("[General]\nportMapping=%d\nlobbyserver=http://127.0.0.1:9\n" % (mode != "off"))
    t0 = time.time()
    p = subprocess.Popen([exe, "--configdir=" + cfg, "--datadir=" + os.path.join(ROOT, "data"), "--vbscript=" + script,
                          "--nosound", "--debug=net", "--debug=console", "--gameport=%d" % GAME_PORT], cwd=os.path.dirname(exe), env=env)
    try:
        p.wait(timeout=240)
        code = p.returncode
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        code = None
    router.close()
    elapsed = time.time() - t0
    logs = sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt")))
    log = open(logs[-1], encoding="utf-8", errors="replace").read() if logs else ""
    open(os.path.join(work, "fake_router_requests.txt"), "w").write("\n".join(router.requests) + "\n")
    for shot in glob.glob(os.path.join(cfg, "screenshots", "*.png")):
        base = os.path.basename(shot)
        shutil.copyfile(shot, os.path.join(work, base.split("-", 2)[2] if base.count("-") >= 2 else base))
    return code, elapsed, log, router.requests


def check(results, name, ok, detail=""):
    results.append((name, bool(ok), detail))


def lines_with(log, *parts):
    return [ln for ln in log.splitlines() if all(p in ln for p in parts)]


def evaluate(mode, code, log, reqs, work):
    r = []

    def P(text):                        # the texts below name port 2100: the run's game port
        return text.replace("2100", str(GAME_PORT))

    def look(name, shot):               # console text: the lobby's chat box shows it (the console logs nothing there)
        r.append((name, None, "look at " + os.path.join(work, shot)))
    leaving = log.find("(VB) script: leaving")
    if leaving < 0:
        leaving = log.find("leaving")
    before, after = (log[:leaving], log[leaving:]) if leaving >= 0 else (log, "")
    check(r, "exe exit code 0", code == 0, "code %s" % code)
    # the safety checks: the override at each lookup, the fake's address as the gateway, nothing else
    overrides = lines_with(log, "(test override) default gateway")
    check(r, "override applied at each gateway lookup", (mode not in ("pcp", "natpmp") or overrides) and
          all(" replaced by 127.0.0.1" in ln for ln in overrides), "%d lookups" % len(overrides))
    check(r, "no override refused", not lines_with(log, "(test override)", "refused"))
    probes = lines_with(log, "gateway address is")
    check(r, "gateway probed is the fake (127.0.0.1:5351)", all("127.0.0.1:5351" in ln for ln in probes),
          "; ".join(ln.split("] ", 1)[-1] for ln in probes[:2]))
    ssdp = lines_with(log, "(test override) SSDP search sent unicast to 127.0.0.1")
    check(r, "SSDP search unicast to the fake", mode in ("pcp", "natpmp", "off") or ssdp, "%d searches" % len(ssdp))
    check(r, "no assert", "assert" not in log.lower() or not lines_with(log, "ASSERT"))
    if mode == "off":
        check(r, "port mapping disabled by user", lines_with(log, "Automatic port mapping setup disabled by user."))
        check(r, "no libplum at all", not lines_with(log, "Plum") and not reqs)
        check(r, "hosting as without port mapping", lines_with(log, "Hosting a server. We are player"))
        return r
    check(r, P("added mapping 0 for port 2100 (callback=1)"), lines_with(log, P("Added mapping 0 for internal port 2100 (callback=1)")))
    check(r, "performing mapping", lines_with(log, P("Performing mapping for internal port 2100")))
    if mode in ("pcp", "natpmp", "upnp"):
        proto = {"pcp": "PCP", "natpmp": "NAT-PMP", "upnp": "UPnP"}[mode]
        if mode == "upnp":
            check(r, "PCP refused by ICMP (10054)", lines_with(log, "UDP recvfrom failed, errno=10054"))
            check(r, "UPnP-IGD found", lines_with(log, "Found UPnP-IGD device: portmap_test/1.0"))
            check(r, "UPnP probe success", lines_with(log, "Success probing UPnP, external address is " + EXTERNAL_IP))
        else:
            check(r, proto + " probe success", lines_with(log, "Success probing %s, gateway address is 127.0.0.1:5351" % proto))
        check(r, "mapping succeeded with the fake's address",
              lines_with(log, P("Mapping succeeded for internal port 2100, external address is %s:2100") % EXTERNAL_IP))
        check(r, "LibPlum cb state=2", lines_with(log, "LibPlum cb: Port mapping 0: state=2"))
        check(r, "Mapping 0: success line", lines_with(log, P("Mapping 0: success, internal=2100, external=%s:2100") % EXTERNAL_IP))
        look(P("console: Port mapping opened external port: 2100"), "pm_hosted.png")
        look("console: Your external IP is: " + EXTERNAL_IP, "pm_hosted.png")
        check(r, "hosting allowed after the mapping", lines_with(log, "Hosting a server. We are player"))
        check(r, "removed mapping when leaving", lines_with(after, P("Removed mapping 0 for internal port 2100")))
        check(r, "NETremRedirects: removed port mapping", lines_with(after, P("Removed port mapping for port 2100 (IPv4).")))
        if mode == "pcp":
            maps = [q for q in reqs if "pcp MAP" in q]
            check(r, P("fake got the MAP (TCP, internal 2100, lifetime 7200)"),
                  any(P("protocol=6 internal=2100") in q and "lifetime(le)=7200" in q for q in maps), "; ".join(maps))
            check(r, "unmapping performed", lines_with(after, P("Unmapped internal port 2100")))
            check(r, "fake got the removal (MAP lifetime 0)", any(P("internal=2100") in q and "lifetime(le)=0 " in q for q in maps))
        elif mode == "natpmp":
            maps = [q for q in reqs if "natpmp MAP" in q]
            check(r, P("fake got the NAT-PMP MAP (TCP, internal 2100, lifetime 7200)"),
                  any("TCP" in q and P("internal=2100") in q and "lifetime=7200" in q for q in maps), "; ".join(maps))
            check(r, "leaving: libplum's NAT-PMP reset (its epoch check), as the official build",
                  lines_with(after, "NAT-PMP reset detected"))
            check(r, "leaving: no unmapping request (the reset dropped the mapping)",
                  not any("lifetime=0" in q for q in maps))
        else:
            posts = [q for q in reqs if "http POST" in q]
            check(r, P("fake got AddPortMapping (2100 TCP, lease 7200)"),
                  any("#AddPortMapping" in q and P("<NewExternalPort>2100</NewExternalPort><NewProtocol>TCP</NewProtocol><NewInternalPort>2100</NewInternalPort>") in q and "<NewLeaseDuration>7200</NewLeaseDuration>" in q
                      for q in posts))
            check(r, "unmapping performed", lines_with(after, P("Unmapped internal port 2100")))
            check(r, P("fake got DeletePortMapping (2100 TCP)"),
                  any("#DeletePortMapping" in q and P("<NewExternalPort>2100</NewExternalPort><NewProtocol>TCP") in q
                      for q in posts))
    elif mode == "fail":
        # the official build's lines (build/interop/official_host's log, hosting with libplum on the real router)
        for text in ("UDP recvfrom failed, errno=10054", "LibPlum cb: Port mapping 0: state=3", "Mapping 0: failed",
                     P("Mapping failed for internal port 2100")):
            check(r, "official line: " + text, lines_with(log, text))
        look("console: Failed to create port mapping", "pm_hosted.png")
        look(P("console: Manually configure your router/firewall to open port 2100!"), "pm_hosted.png")
        check(r, "hosting allowed after the failure", lines_with(log, "Hosting a server. We are player"))
        check(r, "NETremRedirects: removed port mapping", lines_with(after, P("Removed port mapping for port 2100 (IPv4).")))
    elif mode == "timeout":
        look("console: Failed to create port mapping (timeout)", "pm_hosted.png")
        look(P("console: Manually configure your router/firewall to open port 2100!"), "pm_hosted.png")
        check(r, "hosting allowed after the timeout", lines_with(log, "Hosting a server. We are player"))
        check(r, "the IGD never answered AddPortMapping", any("AddPortMapping never answered" in q for q in reqs))
        check(r, "libplum's later failure finds no request (timed out)",
              lines_with(log, "no active request found in PortMappingManager, probably timed out"))
    return r


def main(argv):
    exe = DEFAULT_EXE
    if "--exe" in argv:
        exe = argv[argv.index("--exe") + 1]
        argv = [a for a in argv if a not in ("--exe", exe)]
    modes = argv or ["pcp", "natpmp", "upnp", "fail", "timeout", "off"]
    global GAME_PORT
    for port in range(2150, 2200):      # not 2100: other tests' games (the lobby fork's) host there meanwhile
        t = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
        try:
            t.setsockopt(socket.IPPROTO_IPV6, socket.IPV6_V6ONLY, 0)
            t.bind(("::", port))        # the game's dual-stack listen socket
            GAME_PORT = port
            break
        except OSError:
            pass
        finally:
            t.close()
    print("game port %d" % GAME_PORT)
    os.makedirs(WORK, exist_ok=True)
    out = []
    failed = 0
    for mode in modes:
        code, elapsed, log, reqs = run_scenario(mode, exe)
        res = evaluate(mode, code, log, reqs, os.path.join(WORK, mode))
        out.append("== %s (%.0f s, %d log lines, %d fake router requests)" % (mode, elapsed, len(log.splitlines()), len(reqs)))
        for name, ok, detail in res:
            out.append("  %s %s%s" % ("LOOK" if ok is None else "PASS" if ok else "FAIL", name, (": " + detail) if detail else ""))
            failed += ok is False
        print("\n".join(out[-len(res) - 1:]))
    out.append("%d check(s) failed" % failed)
    print(out[-1])
    open(os.path.join(WORK, "results.txt"), "w").write("\n".join(out) + "\n")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
