"""
lobby_test.py - the internet lobby (cluster J part 2: lib/netplay/netlobby.cpp, src/urlrequest.cpp over WzNet's
urltransfer_c, lib/widget/verifyprompt.cpp) against a local mock lobby: a Python HTTP server on 127.0.0.1 and ::1
(one port, plain HTTP; lobbyserver=http://localhost:<port>) answering what netlobby.cpp asks, recording every
request with the address it came from, and checking them against netlobby.cpp's code, the family included: the
requests 4.7.0 sends over IPv4 only (create, PATCH, DELETE, check-join-request, the tcp:ip4 connection, an ip4
request-join) must come from 127.0.0.1, the tcp:ip6 connection from ::1.

  python -I tools\\lobby_test.py [--exe PATH] [host] [join] [browser] [real]

* host: the game hosts (--autohost: Sk-Rush, an open slot, two NullBots) with lobbyserver=<mock>: the listing's
  lifecycle as the mock sees it: POST /api/v1/game (the game details, joinopts; no attestation: the mock is not
  the trusted lobby), POST .../connection (tcp:ip4 and tcp:ip6 with the host's port), the mock's lobby connection
  check for each (a TCP connection to the host's port over that family: 8 zero bytes, a type-0 message with
  {lobbyToken, challenge}; the host's "WZLR" answer with {"gid": ...}), PATCH heartbeats with the HostToken until the mock says published (the host's
  "Lobby GameId" message), DELETE when the game closes.
* join (with host): a second copy opens Multi Player > Join Game: the browser lists the mock's games (GET
  /api/v1/games, JSON lines); a click on the host's listing asks for join details (POST .../request-join, once
  per IP version); the mock first answers CAPTCHA_REQUIRED with a captcha (a PNG): the verify prompt shows, the
  script types the text and presses Enter; the request comes again with the captcha's response and gets the
  host's address and a join token; the client joins with it and the host checks it (POST .../check-join-request,
  {"pass": true}); the client reaches the host's game lobby.
* browser: only the browser's list (a quick look).
* down: the game hosts with lobbyserver=http://127.0.0.1:9 (nothing answers): the create request fails to connect
  ("Error connecting to the lobby server." in the game's lobby).
* real: ONE read of the real lobby's public games list through the port's transport (HTTPS, the default lobby
  address): the browser opens and shows it; nothing else is sent to the real lobby (no hosting; portMapping=0).
Default: host and join. --exe: the Warzone.exe to run (default bin\\Warzone.exe; a fork's private build, e.g.
build\\j_LB\\bin\\Warzone.exe). Every run's config has portMapping=0 (no router mapping) and, but for real, the mock's
lobby address. Results in build\\lobby_test (screenshots, logs, mock_requests.json).
"""
import base64
import glob
import http.server
import json
import os
import shutil
import socket
import socketserver
import struct
import subprocess
import sys
import threading
import time
import zlib

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
WORK = os.path.join(ROOT, "build", "lobby_test")
GAME_PORT = 2100
HOST_TOKEN = "mock-host-token-1"
LOBBY_TOKEN = "mock-lobby-token-1"
GAME_ID = "mockgame1"
JOIN_TOKEN = "mock-join-token-1"
CAPTCHA_TOKEN = "mock-captcha-token-1"
CAPTCHA_TEXT = "wz2100"
AUTOHOST = """{
	"challenge": {"map": "Sk-Rush", "maxPlayers": 4, "scavengers": "false", "version": 2},
	"player_0": {"team": 0, "name": "LobbyHost"},
	"player_1": {"team": 1},
	"player_2": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 2},
	"player_3": {"ai": "multiplay/skirmish/nb_generic.js", "difficulty": "Medium", "team": 3}
}
"""
# the browser's rows: the first listing (the host's game), the menu's Multi Player and Join Game buttons
MENU_MULTIPLAYER = (512, 400)
MENU_JOIN = (512, 400)
FIRST_ROW = (300, 130)
CAPTCHA_BOX = (492, 357)                # the verify prompt's text box (the joining screen's form, 1024 x 768)


def png_bytes(w, h):
    """A grey captcha-like PNG (RGB, stripes)."""
    raw = bytearray()
    for y in range(h):
        raw.append(0)
        for x in range(w):
            v = 255 if ((x // 8) + (y // 8)) % 2 else 40
            raw += bytes((v, 255 - v, 128))
    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) +
            chunk(b"IDAT", zlib.compress(bytes(raw))) + chunk(b"IEND", b""))


class Mock:
    def __init__(self):
        self.lock = threading.Lock()
        self.requests = []
        self.published = False
        self.check_result = None        # the lobby connection check (tcp:ip4): the host's answer (dict) or an error text
        self.check_result6 = None       # the same over IPv6 (tcp:ip6)
        self.captcha = True             # the first request-join answers CAPTCHA_REQUIRED
        self.server = None

    def record(self, handler, body):
        try:
            parsed = json.loads(body) if body else None
        except ValueError:
            parsed = "(not JSON) " + body[:200]
        entry = {"t": round(time.time(), 2), "method": handler.command, "path": handler.path,
                 "headers": {k.lower(): v for k, v in handler.headers.items()}, "body": parsed,
                 "peer": handler.client_address[0]}
        with self.lock:
            self.requests.append(entry)
        print("  mock: %s %s from %s %s" % (handler.command, handler.path, entry["peer"], (body or "")[:160]))
        return entry

    def games_jsonl(self):
        listing = {"id": GAME_ID, "v": "4.7.0", "netv": [0x4000, 104], "name": "Mock lobby game", "host": {"name": "LobbyHost", "pk": ""},
                   "map": ["Sk-Rush", ""], "players": {"cur": 1, "max": 4}, "specs": {"cur": 0, "max": 2}, "ally": 0, "lim": 0,
                   "conn": ["tcp:ip4"]}
        with self.lock:
            created = [r["body"] for r in self.requests if r["method"] == "POST" and r["path"] == "/api/v1/game" and isinstance(r["body"], dict)]
        if created:
            # the host's own listing, as the lobby lists it (its game details; the joining screen checks the host's key)
            listing.update({k: v for k, v in created[0].items() if k not in ("joinopts", "attest")})
            listing["name"] = listing["name"] or "Mock lobby game"
        other = dict(listing, id="mockgame2", name="Another mock game", players={"cur": 2, "max": 2}, conn=["tcp"])
        bad = {"id": "broken"}          # an invalid listing: logged and skipped ("Invalid listing: ")
        lines = [{"type": "header", "motd": "Mock lobby: hello"}, listing, other, bad]
        return "\n".join(json.dumps(x) for x in lines) + "\n"

    def connection_check(self, port, ip6=False):
        """The lobby's check of a host's connection (netplay.cpp's PendingLobbyConnectCheckRequest)."""
        result = None
        try:
            time.sleep(0.5)
            s = socket.create_connection(("::1" if ip6 else "127.0.0.1", port), timeout=10)
            s.sendall(struct.pack(">II", 0, 0))
            time.sleep(0.5)
            req = json.dumps({"lobbyToken": LOBBY_TOKEN, "challenge": base64.b64encode(os.urandom(32)).decode()}).encode()
            data = struct.pack(">H", len(req)) + req
            # a NetMessage's raw form (netqueue.cpp): type (0), the payload's length (uint16, network order), payload
            s.sendall(bytes([0]) + struct.pack(">H", len(data)) + data)
            buf = b""
            while len(buf) < 12:
                d = s.recv(4096)
                if not d:
                    break
                buf += d
            if len(buf) >= 12 and buf[:4] == b"WZLR":
                ver, n = struct.unpack(">II", buf[4:12])
                while len(buf) < 12 + n:
                    d = s.recv(4096)
                    if not d:
                        break
                    buf += d
                result = {"version": ver, "response": json.loads(buf[12:12 + n].decode())}
            else:
                result = "no WZLR answer: %r" % buf[:40]
            s.close()
        except Exception as e:      # noqa: BLE001 (a test's report)
            result = "connection check failed: %s" % e
        if ip6:
            self.check_result6 = result
        else:
            self.check_result = result
        print("  mock: connection check (%s): %s" % ("ip6" if ip6 else "ip4", result))

    def status(self):
        ok = isinstance(self.check_result, dict) and self.check_result.get("response", {}).get("gid") == GAME_ID
        self.published = self.published or ok
        return {"gameId": GAME_ID, "published": self.published,
                "connections": [{"type": "tcp:ip4", "status": "validated" if ok else "submitted"}]}


MOCK = Mock()


class Handler(http.server.BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass

    def reply(self, code, obj=None, text=None, ctype="application/json"):
        body = (text if text is not None else (json.dumps(obj) if obj is not None else "")).encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n).decode("utf-8", "replace") if n else ""

    def do_GET(self):
        MOCK.record(self, "")
        if self.path == "/api/v1/games":
            self.reply(200, text=MOCK.games_jsonl(), ctype="application/jsonl")
        else:
            self.reply(404, {"code": "NOT_FOUND", "message": "no such path"})

    def do_POST(self):
        b = self.body()
        e = MOCK.record(self, b)
        p = self.path
        if p == "/api/v1/game":
            self.reply(201, {"gameId": GAME_ID, "hostToken": HOST_TOKEN, "lobbyToken": LOBBY_TOKEN, "motd": "Mock lobby: game created",
                             "hostJoinOpts": ["botProt", "proxyIps", "unknownOpt"]})
        elif p == "/api/v1/game/%s/connection" % GAME_ID:
            port = (e["body"] or {}).get("port", GAME_PORT)
            ip6 = (e["body"] or {}).get("type") == "tcp:ip6"
            threading.Thread(target=MOCK.connection_check, args=(port, ip6), daemon=True).start()
            self.reply(202, MOCK.status())
        elif p == "/api/v1/game/%s/check-join-request" % GAME_ID:
            self.reply(200, {"pass": (e["body"] or {}).get("joinToken") == JOIN_TOKEN})
        elif p.endswith("/request-join"):
            if MOCK.captcha and "captcha" not in (e["body"] or {}):
                img = base64.b64encode(png_bytes(160, 48)).decode()
                self.reply(403, {"code": "CAPTCHA_REQUIRED", "message": "Please verify",
                                 "data": {"type": "captcha", "token": CAPTCHA_TOKEN, "data": {"image": {"data": img, "type": "png"}}}})
            else:
                self.reply(200, {"connections": [{"type": "tcp:ip4", "host": "127.0.0.1", "port": GAME_PORT}], "joinToken": JOIN_TOKEN})
        else:
            self.reply(404, {"code": "NOT_FOUND", "message": "no such path"})

    def do_PATCH(self):
        MOCK.record(self, self.body())
        if self.path == "/api/v1/game/%s" % GAME_ID:
            self.reply(200, MOCK.status())
        else:
            self.reply(404, {"code": "GAME_NOT_FOUND", "message": "no such game"})

    def do_DELETE(self):
        MOCK.record(self, self.body())
        self.reply(200, {})


class Server(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = False

    def handle_error(self, request, client_address):
        pass                            # a keep-alive connection the game resets when it closes


class Server6(Server):
    address_family = socket.AF_INET6

    def server_bind(self):
        self.socket.setsockopt(socket.IPPROTO_IPV6, socket.IPV6_V6ONLY, 1)
        super().server_bind()


def make_config(name, lobby, autohost=False):
    cfg = os.path.join(WORK, name, "config")
    if os.path.exists(cfg):
        shutil.rmtree(cfg)
    os.makedirs(cfg)
    with open(os.path.join(cfg, "config"), "w", newline="\r\n") as f:
        # never a port mapping on the router; the lobby: the mock (or, for real, 4.7.0's default)
        f.write("[General]\nportMapping=0\n" + ("lobbyserver=%s\n" % lobby if lobby else ""))
    if autohost:
        os.makedirs(os.path.join(cfg, "autohost"))
        with open(os.path.join(cfg, "autohost", "lobby_ai.json"), "w") as f:
            f.write(AUTOHOST)
    return cfg


def script(name, lines):
    p = os.path.join(WORK, name, name + ".txt")
    with open(p, "w") as f:
        f.write("\n".join(lines) + "\n")
    return p


def run_game(exe, name, cfg, scr, extra, limit):
    args = [exe, "--configdir=" + cfg, "--vbscript=" + scr, "--nosound", "--datadir=" + os.path.join(ROOT, "data")] + extra
    return subprocess.Popen(args, cwd=os.path.dirname(exe))


def finish(p, limit, t0):
    try:
        p.wait(timeout=max(1, limit - (time.time() - t0)))
    except subprocess.TimeoutExpired:
        p.kill()
        p.wait()
        return "killed at the time limit"
    return "exit code %d" % p.returncode


def log_lines(cfg, keys):
    out = []
    for f in sorted(glob.glob(os.path.join(cfg, "logs", "WZlog-*.txt"))):
        for ln in open(f, encoding="utf-8", errors="replace"):
            if any(k.lower() in ln.lower() for k in keys):
                out.append(ln.rstrip()[:220])
    return out


def copy_shots(cfg, name):
    got = []
    for shot in sorted(glob.glob(os.path.join(cfg, "screenshots", "*.png"))):
        base = os.path.basename(shot)
        out = os.path.join(WORK, name, (base.split("-", 2)[2] if base.count("-") >= 2 else base))
        shutil.copyfile(shot, out)
        got.append(out)
    return got


CHECKS = []


def check(ok, what):
    CHECKS.append((bool(ok), what))
    print("  [%s] %s" % ("ok" if ok else "FAIL", what))


def reqs(method, path_end=""):
    return [r for r in MOCK.requests if r["method"] == method and r["path"].endswith(path_end)]


def from_v4(rs):
    return bool(rs) and all(r["peer"] == "127.0.0.1" for r in rs)


def check_common_headers(r, referer_tail, accept):
    h = r["headers"]
    return (h.get("origin") == "wz2100://wz2100.net" and h.get("referer", "").endswith(referer_tail) and h.get("accept") == accept
            and h.get("user-agent", "").startswith("Warzone2100/"))


def main(argv):
    exe = os.path.join(ROOT, "bin", "Warzone.exe")
    if "--exe" in argv:
        exe = os.path.abspath(argv[argv.index("--exe") + 1])
    modes = [a for a in argv if a in ("host", "join", "browser", "real", "down")] or ["host", "join"]
    os.makedirs(WORK, exist_ok=True)
    srv = Server(("127.0.0.1", 0), Handler)
    port = srv.server_address[1]
    srv6 = Server6(("::1", port), Handler)
    lobby = "http://localhost:%d" % port
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    threading.Thread(target=srv6.serve_forever, daemon=True).start()
    print("mock lobby at %s; exe %s" % (lobby, exe))

    procs = []
    if "host" in modes:
        cfg = make_config("host", lobby, autohost=True)
        hold = 5400 if "join" in modes else 2700       # frames in the host's lobby (the client joins meanwhile)
        scr = script("host", ["wait 900", "shot host_lobby", "wait %d" % hold, "shot host_lobby2", "log done", "wait 5", "quit"])
        t0 = time.time()
        procs.append(("host", cfg, run_game(exe, "host", cfg, scr, ["--autohost=lobby_ai.json", "--debug=net"], 300), t0, 300))
    if "join" in modes or "browser" in modes:
        if "host" in modes:
            # the client asks for join details once the host's listing exists
            for _ in range(120):
                if reqs("POST", "/api/v1/game"):
                    break
                time.sleep(1)
            time.sleep(3)
        name = "join" if "join" in modes else "browser"
        cfg = make_config(name, lobby)
        lines = ["wait 600", "click %d %d" % MENU_MULTIPLAYER, "wait 60", "click %d %d" % MENU_JOIN, "wait 300", "shot browser", "widgets"]
        if name == "join":
            lines += ["click %d %d" % FIRST_ROW, "wait 300", "shot captcha", "click %d %d" % CAPTCHA_BOX, "wait 10", "text " + CAPTCHA_TEXT,
                      "wait 10", "key 13",
                      "wait 600", "shot joined", "wait 300", "shot joined2"]
        lines += ["log done", "wait 5", "quit"]
        scr = script(name, lines)
        t0 = time.time()
        procs.append((name, cfg, run_game(exe, name, cfg, scr, [], 240), t0, 240))
    if "down" in modes and "host" in modes:
        modes.remove("down")
        print("note: down hosts on port %d as host does: run it alone (python -I tools\\lobby_test.py down)" % GAME_PORT)
    if "down" in modes:
        # hosting while the lobby cannot be reached (ui_test.py's default lobbyserver=http://127.0.0.1:9)
        cfg = make_config("down", "http://127.0.0.1:9", autohost=True)
        scr = script("down", ["wait 1200", "shot down_lobby", "log done", "wait 5", "quit"])
        t0 = time.time()
        procs.append(("down", cfg, run_game(exe, "down", cfg, scr, ["--autohost=lobby_ai.json", "--debug=net"], 150), t0, 150))
    if "real" in modes:
        cfg = make_config("real", None)
        scr = script("real", ["wait 600", "click %d %d" % MENU_MULTIPLAYER, "wait 60", "click %d %d" % MENU_JOIN, "wait 600",
                              "shot real_browser", "log done", "wait 5", "quit"])
        t0 = time.time()
        procs.append(("real", cfg, run_game(exe, "real", cfg, scr, [], 120), t0, 120))

    # the client first (it quits first), then the host
    for name, cfg, p, t0, limit in sorted(procs, key=lambda x: x[0] == "host"):
        print("%s: %s" % (name, finish(p, limit, t0)))
        for s in copy_shots(cfg, name):
            print("  screenshot: %s" % s)
        for ln in log_lines(cfg, ("lobby", "cURL", "urlTransfer", "lobbyerror", "lobbyid", "error", "assert", "captcha", "(VB) script")):
            if "khr_callback" not in ln:
                print("  | " + ln)
    srv.shutdown()
    srv6.shutdown()
    with open(os.path.join(WORK, "mock_requests.json"), "w") as f:
        json.dump(MOCK.requests, f, indent=1)

    print("== checks")
    if "host" in modes:
        c = reqs("POST", "/api/v1/game")
        check(len(c) == 1, "one create request (POST /api/v1/game)")
        check(from_v4(c), "create over IPv4 (InternetProtocol::IPv4): %s" % [r["peer"] for r in c])
        if c:
            b = c[0]["body"] or {}
            check(check_common_headers(c[0], "?wz_internal=mp.host", "application/json") and c[0]["headers"].get("content-type") == "application/json",
                  "create's headers (Origin, Referer mp.host, Accept, Content-Type, User-Agent)")
            check(isinstance(b.get("name"), str) and set(b.get("host", {})) == {"name", "pk"} and len(base64.b64decode(b["host"]["pk"])) == 32
                  and b.get("netv") == [0x4000, 104] and b.get("map", [None])[0] == "Sk-Rush" and b.get("players", {}).get("max") == 4
                  and "specs" in b and "ally" in b and "lim" in b,
                  "create's game details (v, netv, name, host, map, players, specs, ally, lim)")
            check(b.get("joinopts") == {"botProt": "block", "proxyIps": "blockSus", "hostingIps": "block"}, "create's joinopts (4.7.0's defaults)")
            check("attest" not in b, "no attestation for an untrusted lobby")
            check(list(b.keys())[:2] == ["v", "netv"], "the body's key order (ordered_json): v, netv, ...")
        conn = reqs("POST", "/connection")
        c4 = [r for r in conn if (r["body"] or {}).get("type") == "tcp:ip4"]
        c6 = [r for r in conn if (r["body"] or {}).get("type") == "tcp:ip6"]
        check(from_v4(c4) and c6 and all(r["peer"] == "::1" for r in c6),
              "the connection requests over their own family (tcp:ip4 from 127.0.0.1, tcp:ip6 from ::1): %s"
              % [((r["body"] or {}).get("type"), r["peer"]) for r in conn])
        check(len(c4) == 1 and c4[0]["body"] == {"type": "tcp:ip4", "host": "", "port": GAME_PORT}
              and all(r["headers"].get("authorization") == "HostToken " + HOST_TOKEN for r in conn),
              "the connection requests (tcp:ip4 with port %d; %s; HostToken)" % (GAME_PORT, [(r["body"] or {}).get("type") for r in conn]))
        cr = MOCK.check_result
        check(isinstance(cr, dict) and cr.get("version") == 2 and cr.get("response") == {"gid": GAME_ID},
              "the host's answer to the lobby connection check (WZLR, version 2, {\"gid\": ...}): %s" % (cr,))
        patches = reqs("PATCH", "/api/v1/game/" + GAME_ID)
        check(len(patches) >= 1 and all(p["headers"].get("authorization") == "HostToken " + HOST_TOKEN for p in patches),
              "listing updates (PATCH with the HostToken): %d" % len(patches))
        check(from_v4(patches), "the updates over IPv4")
        cr6 = MOCK.check_result6
        check(isinstance(cr6, dict) and cr6.get("response") == {"gid": GAME_ID},
              "the host's answer to the lobby connection check over IPv6 (::1): %s" % (cr6,))
        check(MOCK.published, "the mock published the game")
        dels = reqs("DELETE", "/api/v1/game/" + GAME_ID)
        check(len(dels) == 1 and dels[0]["headers"].get("authorization") == "HostToken " + HOST_TOKEN and from_v4(dels),
              "the listing removed (DELETE, over IPv4) when the host closed")
        hl = "\n".join(log_lines(os.path.join(WORK, "host", "config"), ("lobby game id", "using game id")))
        check("Created lobby game ID: %s - awaiting publish" % GAME_ID in hl and "Using game ID: %s" % GAME_ID in hl,
              "the host's log: created, then published")
    if "join" in modes or "browser" in modes:
        g = reqs("GET", "/api/v1/games")
        check(len(g) >= 1 and check_common_headers(g[0], "?wz_internal=list", "application/jsonl"),
              "the browser's games list request (GET /api/v1/games, IP_ANY: from %s)" % [r["peer"] for r in g])
    if "join" in modes:
        rj = reqs("POST", "/api/v1/game/%s/request-join" % GAME_ID)
        first = [r for r in rj if "captcha" not in (r["body"] or {})]
        second = [r for r in rj if "captcha" in (r["body"] or {})]
        check(len(first) >= 1 and all(r["body"] == {"asSpec": False} for r in first) and check_common_headers(first[0], "?wz_internal=request.join", "application/json"),
              "request-join (the listing's IP version: %d request), {\"asSpec\": false}" % len(first))
        check(from_v4(rj), "request-join for the tcp:ip4 listing over IPv4: %s" % [r["peer"] for r in rj])
        check(len(second) >= 1 and second[0]["body"]["captcha"] == {"response": {"result": CAPTCHA_TEXT}, "token": CAPTCHA_TOKEN},
              "request-join again with the captcha's response (the verify prompt)")
        if "host" in modes:
            cj = reqs("POST", "/check-join-request")
            check(len(cj) >= 1 and (cj[0]["body"] or {}).get("joinToken") == JOIN_TOKEN and (cj[0]["body"] or {}).get("ip") == "127.0.0.1"
                  and (cj[0]["body"] or {}).get("connMethod") == "tcp" and from_v4(cj),
                  "the host checked the join token (check-join-request, over IPv4)")
    if "down" in modes:
        dl = "\n".join(log_lines(os.path.join(WORK, "down", "config"), ("cURL",)))
        check("cURL: Couldn't connect for (http://127.0.0.1:9/api/v1/game)" in dl, "an unreachable lobby: the create request fails (COULDNT_CONNECT)")
    failed = [w for ok, w in CHECKS if not ok]
    print("%d of %d checks passed" % (len(CHECKS) - len(failed), len(CHECKS)))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
