"""
net_test.py - cluster J's test of WzNet's connection layer (fork JT: lib/netplay's client_connection, the provider
registry, error_categories, ip_helpers, the listen socket, the poll group, polling_util, the select() descriptor
set, the pending writes, the zlib adapter, wz_connection_provider, tcp/*): builds build/net_test/NetTests.exe
(those modules with their classes made private, src/shared, the zlib streams of src/shared/miniz, and
src/nettests: a stand-in for WzFramework's log that writes into the results), runs it once per scenario and
checks its results, with a Python peer on the other end where the scenario has one:

* loopback, plain and compressed: listen, connect (127.0.0.1), accept; 200,000 bytes each way in uneven chunks,
  each written with writeAll and flush, read as netplay reads (a poll group, readNoInt when readReady), and the
  same with 7-byte reads; readAll; connectionStatus; the peer closing (ECONNRESET);
* a Python peer: the port connects and sends (plain, and zlib-compressed: Python's zlib inflates the stream,
  byte for byte), the port connects and receives (Python's zlib deflates with Z_SYNC_FLUSH and Z_PARTIAL_FLUSH
  at several levels), the port listens and Python connects over IPv4 and IPv6 (the dual-stack listen socket);
* openClientConnectionAsync: the callback runs before the call returns; a refused connection and an
  unreachable address (timed out), with their error texts; resolveHost (localhost, and a name that does not
  resolve: getaddrinfo's message, FormatMessageA's text as Windows' gai_strerror gives it);
* the error codes: GenericSystemErrorCategory's messages, MSVC's generic messages, the zlib category's, and
  `ec == std::errc::x` as MSVC compares (a Winsock code maps, an errno value from client_connection does not);
* readAll's timeout (MSVC's ETIMEDOUT, 138: "Unknown error", not errc::timed_out on Windows);
* the pending writes under a full socket buffer: 8 MB queued at once while the peer does not read, drained by
  PendingWritesManager_pollAll passes, all of it in order;
* the byte order and IP helpers.

  python -I tools\\net_test.py [--keep] [--stage DIR] [--zlib DIR]
      --stage: the folder holding net\\ and nettests\\ (default: src\\); --zlib: the zlib streams' folder (default:
      src\\shared\\miniz; a stand-in folder skips the compressed scenarios)
"""
import ctypes
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import zlib

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

WORK = os.path.join(ROOT, "build", "net_test")
# the connection layer's modules (cluster J, fork JT), in src/net
JT_MODULES = ["byteorder_funcs_wrapper_c.bas", "connection_provider_registry_c.bas", "error_categories_c.bas",
              "ip_helpers_c.bas", "netsocket_c.bas", "pending_writes_manager_map_c.bas", "polling_util_c.bas",
              "wz_compression_provider_c.bas", "ConnectionWriteQueue.cls", "IClientConnection.cls",
              "IConnectionAddress.cls", "IConnectionPollGroup.cls", "IListenSocket.cls", "IOpenConnectionCallback.cls",
              "OpenConnectionResult.cls", "PendingWritesManager.cls", "SelectDescriptorSet.cls",
              "TCPAddressResolver.cls", "TCPClientConnection.cls", "TCPConnectionAddress.cls",
              "TCPConnectionPollGroup.cls", "TCPConnectionProvider.cls", "TCPListenSocket.cls",
              "WzConnectionProvider.cls", "ZlibCompressionAdapter.cls"]
PORT_BASE = 30500


def gen(seed, n):
    """nettests_main's gen(seed, i) for i < n."""
    return bytes(((i * 31 + seed * 7 + (i >> 8) + (i >> 16) * 3) & 0xFF) for i in range(n))


def build_exe(stage, zlib_dir):
    src = os.path.join(WORK, "src")
    if os.path.exists(src):
        shutil.rmtree(src)
    os.makedirs(src)
    files = [os.path.join(stage, "net", n) for n in JT_MODULES]
    files += [os.path.join(stage, "nettests", n) for n in build.sources(os.path.join(stage, "nettests"))]
    files += [os.path.join(zlib_dir, n) for n in build.sources(zlib_dir)]
    files += [os.path.join(build.SHARED, n) for n in build.sources(build.SHARED)]
    lines = ["Type=Exe", build.STDOLE]
    for f in files:
        dst = os.path.join(src, os.path.basename(f))
        data = open(f, "rb").read().decode("latin-1").replace("\r\n", "\n")
        # a Standard EXE holds private classes only: WzNet's public ones become private here
        data = data.replace("Attribute VB_Exposed = True", "Attribute VB_Exposed = False")
        data = data.replace("Attribute VB_Creatable = False", "Attribute VB_Creatable = True")
        open(dst, "wb").write(data.replace("\n", "\r\n").encode("latin-1"))
        name = build.vb_name(dst)
        if dst.lower().endswith(".cls"):
            lines.append("Class=%s; %s" % (name, os.path.basename(dst)))
        else:
            lines.append("Module=%s; %s" % (name, os.path.basename(dst)))
    lines.append(build.VBP_TAIL.format(name="NetTests", ext="exe", path=".", startup="Sub Main", compat="",
                                       bounds="0", optimisation="2", threading=""))
    vbp = os.path.join(src, "NetTests.vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    log = os.path.join(WORK, "NetTests_build.log")
    if os.path.exists(log):
        os.remove(log)
    subprocess.run([build.VB6, "/make", vbp, "/outdir", WORK, "/out", log], cwd=src)
    text = open(log, encoding="latin-1").read().strip() if os.path.exists(log) else "VB6 stopped without a message"
    print(text)
    return "succeeded" in text


def run_cmd(cmd, timeout=120):
    """NetTests.exe on one command: (results dict, log lines, assert lines)."""
    cf = os.path.join(WORK, "cmd.txt")
    rf = os.path.join(WORK, "results.txt")
    if os.path.exists(rf):
        os.remove(rf)
    with open(cf, "w") as f:
        f.write("\t".join(str(c) for c in cmd) + "\n")
    subprocess.run([os.path.join(WORK, "NetTests.exe"), cf, rf], timeout=timeout)
    results, logs, asserts = {}, [], []
    if os.path.exists(rf):
        for line in open(rf, encoding="latin-1").read().splitlines():
            parts = line.split("\t")
            if parts[0] == "result":
                results[parts[1]] = parts[2] if len(parts) > 2 else ""
            elif parts[0] == "log":
                logs.append(parts)
            elif parts[0] == "assert":
                asserts.append(parts)
    return results, logs, asserts


class Peer(threading.Thread):
    """A Python end: fn(sock) runs on the accepted (server) or connected (client) socket."""

    def __init__(self, fn, port=None, connect=None, family=socket.AF_INET):
        super().__init__(daemon=True)
        self.fn, self.port, self.connect, self.family = fn, port, connect, family
        self.result, self.error = None, None
        if connect is None:
            self.srv = socket.socket(family, socket.SOCK_STREAM)
            self.srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.srv.bind(("127.0.0.1" if family == socket.AF_INET else "::1", port))
            self.srv.listen(1)
            self.srv.settimeout(30)

    def run(self):
        try:
            if self.connect is None:
                s, _ = self.srv.accept()
            else:
                t0 = time.time()
                while True:
                    try:
                        s = socket.create_connection(self.connect, timeout=5)
                        break
                    except OSError:
                        if time.time() - t0 > 15:
                            raise
                        time.sleep(0.05)
            s.settimeout(30)
            self.result = self.fn(s)
            s.close()
        except Exception as e:                     # noqa: BLE001 (reported)
            self.error = repr(e)
        finally:
            if self.connect is None:
                self.srv.close()


def recv_all(s):
    out = bytearray()
    while True:
        b = s.recv(65536)
        if not b:
            return bytes(out)
        out += b


class Checker:
    def __init__(self):
        self.ok = self.total = 0
        self.failed = []

    def check(self, name, cond, detail=""):
        self.total += 1
        if cond:
            self.ok += 1
        else:
            self.failed.append(name)
            print("  FAIL %s %s" % (name, detail))
        return cond


def format_message(code):
    """Windows' gai_strerrorA: FormatMessageA(FROM_SYSTEM | IGNORE_INSERTS | MAX_WIDTH_MASK, ..., LANG_NEUTRAL)."""
    buf = ctypes.create_string_buffer(1025)
    n = ctypes.windll.kernel32.FormatMessageA(0x1000 | 0x200 | 0xFF, None, code, 0x400, buf, 1024, None)
    return buf.raw[:n].decode("mbcs")


def main(argv):
    keep = "--keep" in argv
    stage = os.path.join(ROOT, "src")
    zlib_dir = os.path.join(build.SHARED, "miniz")
    if "--stage" in argv:
        stage = argv[argv.index("--stage") + 1]
    if "--zlib" in argv:
        zlib_dir = argv[argv.index("--zlib") + 1]
    real_zlib = any(n.lower().startswith("miniz") for n in build.sources(zlib_dir))
    os.makedirs(WORK, exist_ok=True)
    if not build_exe(stage, zlib_dir):
        return 1
    c = Checker()
    asserts_total = 0
    port = PORT_BASE

    def scenario(name, cmd, timeout=120):
        nonlocal asserts_total
        r, logs, asserts = run_cmd(cmd, timeout)
        asserts_total += len(asserts)
        for a in asserts:
            print("  assert in %s: %s" % (name, " | ".join(a[1:])))
        c.check(name + ": ran", r.get("done") == "1", str(r))
        return r, logs

    # helpers
    r, _ = scenario("helpers", ["helpers"])
    c.check("htonl", r.get("htonl") == "78563412" and r.get("htonl2") == "201FF80" and r.get("ntohl") == "12345678", str(r))
    c.check("htons", r.get("htons") == "3412" and r.get("htons2") == "80FF", str(r))
    c.check("unaligned", r.get("store") == "AB CD" and r.get("load") == "ABCD", str(r))
    expect_class = {"10.1.2.3": 0, "172.16.0.1": 0, "172.32.0.1": 3, "192.168.1.1": 0, "169.254.3.4": 1,
                    "127.0.0.1": 2, "8.8.8.8": 3, "::1": 2, "fe80::1": 1, "fd00::5": 0, "2001:db8::1": 3,
                    "::ffff:192.168.0.9": 0, "::ffff:8.8.4.4": 3, "nonsense": -1, "": -1}
    expect_format = {"10.1.2.3": 0, "::1": 1, "fe80::1": 1, "::ffff:192.168.0.9": 2, "nonsense": -1, "": -1}
    c.check("ip classification", all(r.get("class_" + k) == str(v) for k, v in expect_class.items()) and
            r.get("class_null") == "-1", str({k: r.get("class_" + k) for k in expect_class}))
    c.check("ip format", all(r.get("format_" + k) == str(v) for k, v in expect_format.items()) and
            r.get("format_names") == "ip4,ip6,ip4in6,", str(r.get("format_names")))

    # error codes
    r, _ = scenario("errc", ["errc"])
    gs = {0: "No error", 9: "Unknown error", 14: "Unknown error", 22: "Unknown error", 108: "Unknown error",
          138: "Unknown error", 10004: "Interrupted system call", 10009: "Bad file number",
          10035: "Operation would block", 10036: "Operation now in progress", 10054: "Connection reset by peer",
          10060: "Connection timed out", 10061: "Connection refused", 10065: "No route to host",
          11001: "Host not found", 11004: "No host data of that type was found", 12345: "Unknown error"}
    c.check("generic system messages", all(r.get("gs_%d" % k) == v for k, v in gs.items()),
            str({k: r.get("gs_%d" % k) for k in gs}))
    gen_msgs = {4: "interrupted", 11: "resource unavailable try again", 106: "connection aborted",
                107: "connection refused", 108: "connection reset", 138: "timed out", 140: "operation would block",
                999: "unknown error"}
    c.check("generic messages", all(r.get("generic_%d" % k) == v for k, v in gen_msgs.items()),
            str({k: r.get("generic_%d" % k) for k in gen_msgs}))
    c.check("getaddrinfo messages (FormatMessageA)", all(r.get("gai_%d" % k) == format_message(k)
                                                         for k in (11001, 11002, 11003, 10109)),
            str({k: (r.get("gai_%d" % k), format_message(k)) for k in (11001, 10109)}))
    c.check("zlib messages", r.get("zlib_2") == "Z_NEED_DICT" and r.get("zlib_-2") == "Z_STREAM_ERROR" and
            r.get("zlib_-3") == "Z_DATA_ERROR" and r.get("zlib_-4") == "Z_MEM_ERROR" and
            r.get("zlib_-5") == "Unknown zlib error", str(r))
    eqs = {"eq_10060_timed_out": "True", "eq_138_timed_out": "False", "eq_10035_would_block": "True",
           "eq_10035_try_again": "False", "eq_10054_reset": "True", "eq_108_reset": "False",
           "eq_10004_interrupted": "True", "eq_generic_106_aborted": "True", "eq_gai_11001_timed_out": "False",
           "eq_zero": "False", "value_zlib_-3": "-3", "value_0_is_error": "True"}
    c.check("errc comparisons (MSVC's)", all(r.get(k) == v for k, v in eqs.items()), str({k: r.get(k) for k in eqs}))

    # loopback in the exe
    for name, compressed, n, chunk, read in (("loopback plain", 0, 200000, 3000, 1500),
                                             ("loopback plain 7-byte reads", 0, 20000, 300, 7),
                                             ("loopback compressed", 1, 200000, 3000, 1500),
                                             ("loopback compressed 7-byte reads", 1, 20000, 300, 7)):
        if compressed and not real_zlib:
            print("  skipped %s (no zlib streams)" % name)
            continue
        port += 1
        r, logs = scenario(name, ["loopback", port, n, compressed, chunk, read])
        c.check(name + ": data both ways", r.get("got1") == str(n) and r.get("bad1") == "0" and
                r.get("got2") == str(n) and r.get("bad2") == "0" and r.get("recv1_ec") == "0" and
                r.get("recv2_ec") == "0", str(r))
        c.check(name + ": one dual-stack listen socket", any("Successfully enabled IPv4 to IPv6 mapping" in l[3]
                                                             for l in logs), str([l[3] for l in logs][:6]))
        # (connectionStatus on a compressed connection whose inflater still holds data reports ECONNRESET: 4.7.0's
        # checkConnectionsReadable marks it ready without polling, and its socket's read queue is empty)
        c.check(name + ": addresses, status", r.get("clientAddress") == "127.0.0.1" and
                r.get("serverAddress") == "127.0.0.1" and r.get("supportedIpVersions") == "3" and
                (compressed or (r.get("status_client_ec") == "0" and r.get("status_server_ec") == "0")) and
                r.get("compressed") == ("True" if compressed else "False"), str(r))
        if not compressed:
            c.check(name + ": readAll", r.get("readAll_ec") == "0" and r.get("readAll_count") == "16" and
                    r.get("readAll_bad") == "0", str(r))
        c.check(name + ": the peer closing", r.get("status_after_close_message") == "Connection reset by peer" and
                r.get("read_after_close_message") == "Connection reset by peer" and
                r.get("clientValidAfterClose") == "False", str(r))
        print("  %s: %s ms" % (name, r.get("ms")))

    # the port sends to Python
    for name, compressed in (("connect_send plain", 0), ("connect_send compressed", 1)):
        if compressed and not real_zlib:
            continue
        port += 1
        n = 300000

        def peer_recv(s, compressed=compressed):
            data = recv_all(s)
            if compressed:
                d = zlib.decompressobj()
                return d.decompress(data) + d.flush(), len(data)
            return data, len(data)

        p = Peer(peer_recv, port=port)
        p.start()
        r, _ = scenario(name, ["connect_send", "127.0.0.1", port, n, 5, compressed, 4000])
        p.join(60)
        got, raw = p.result if p.result else (b"", 0)
        c.check(name + ": Python got it", got == gen(5, n) and r.get("drained") == "True",
                "%s %d/%d %s" % (p.error, len(got), n, r))
        if compressed:
            print("  %s: %d bytes compressed to %d" % (name, n, raw))

    # Python sends to the port
    for name, compressed, level, flush in (("connect_recv plain", 0, 0, 0),
                                           ("connect_recv compressed sync flush", 1, 6, zlib.Z_SYNC_FLUSH),
                                           ("connect_recv compressed partial flush level 1", 1, 1, zlib.Z_PARTIAL_FLUSH),
                                           ("connect_recv compressed partial flush level 9", 1, 9, zlib.Z_PARTIAL_FLUSH)):
        if compressed and not real_zlib:
            continue
        port += 1
        n = 250000

        def peer_send(s, compressed=compressed, level=level, flush=flush):
            assert s.recv(1) == b"G"
            data = gen(9, n)
            if compressed:
                co = zlib.compressobj(level)
                pos, k = 0, 0
                while pos < n:
                    k += 1
                    chunk = (k * 4099) % 5000 + 1
                    s.sendall(co.compress(data[pos:pos + chunk]) + co.flush(flush))
                    pos += chunk
            else:
                s.sendall(data)
            time.sleep(1)
            return True

        p = Peer(peer_send, port=port)
        p.start()
        r, _ = scenario(name, ["connect_recv", "127.0.0.1", port, n, 9, compressed, 2000])
        p.join(60)
        c.check(name + ": the port got it", r.get("got") == str(n) and r.get("bad") == "0" and
                r.get("recv_ec") == "0", "%s %s" % (p.error, r))

    # the port listens, Python connects (IPv4, and IPv6 to the dual-stack socket)
    for name, host, family, expect_addr in (("listen_recv IPv4", "127.0.0.1", socket.AF_INET, "127.0.0.1"),
                                            ("listen_recv IPv6", "::1", socket.AF_INET6, "0:0:0:0:0:0:0:1")):
        port += 1
        n = 100000

        def peer_conn(s):
            s.sendall(gen(4, n))
            time.sleep(1)
            return True

        p = Peer(peer_conn, connect=(host, port), family=family)
        p.start()
        r, _ = scenario(name, ["listen_recv", port, n, 4, 4096])
        p.join(60)
        c.check(name, r.get("got") == str(n) and r.get("bad") == "0" and r.get("address") == expect_addr,
                "%s %s" % (p.error, r))

    # openClientConnectionAsync
    port += 1
    p = Peer(lambda s: True, port=port)
    p.start()
    r, _ = scenario("async connected", ["async", "127.0.0.1", port, 3000])
    p.join(10)
    c.check("async connected: callback before return", r.get("calledBeforeReturn") == "1" and
            r.get("returned") == "True" and r.get("hasError") == "False" and r.get("address") == "127.0.0.1", str(r))
    port += 1
    r, logs = scenario("async refused", ["async", "127.0.0.1", port, 3000])
    msg = r.get("error_message")
    c.check("async refused: an error with its text", r.get("calledBeforeReturn") == "1" and r.get("hasError") == "True"
            and r.get("errorString") == 'Cannot resolve host "127.0.0.1": [%s]: %s' % (r.get("error_value"), msg),
            str(r))
    print("  refused: [%s] %s (%s)" % (r.get("error_value"), msg, [l[3] for l in logs if "socketOpen" in l[2]][-1:]))
    r, logs = scenario("async unreachable", ["async", "10.255.255.1", 9, 700])
    print("  unreachable: [%s] %s in %s ms" % (r.get("error_value"), r.get("error_message"), r.get("ms")))
    c.check("async unreachable: an error", r.get("hasError") == "True", str(r))
    if r.get("error_value") == "10060":
        c.check("async unreachable: timed out", r.get("error_message") == "Connection timed out" and
                r.get("timed_out") == "True", str(r))

    # resolveHost
    r, _ = scenario("resolve localhost", ["resolve", "localhost", 2100])
    c.check("resolve localhost", r.get("resolve_ec") == "0" and r.get("address") in ("[::1]:2100", "127.0.0.1:2100"),
            str(r))
    r, logs = scenario("resolve unknown", ["resolve", "nonexistent.invalid", 2100])
    gai_value = r.get("resolve_value")
    c.check("resolve unknown: getaddrinfo's error", r.get("resolve_categoryName") == "getaddrinfo" and gai_value and
            r.get("resolve_message") == format_message(int(gai_value)) and
            any("getaddrinfo failed for nonexistent.invalid:2100: " in l[3] for l in logs), str(r))

    # readAll's timeout
    port += 1
    r, _ = scenario("readAll timeout", ["readall_timeout", port])
    c.check("readAll timeout: MSVC's ETIMEDOUT", r.get("readAll_value") == "138" and
            r.get("readAll_message") == "Unknown error" and r.get("timed_out") == "False", str(r))

    # the pending writes under a full socket buffer
    port += 1
    n = 8 * 1024 * 1024

    def peer_slow(s):
        time.sleep(1.5)
        return recv_all(s)

    p = Peer(peer_slow, port=port)
    p.start()
    r, _ = scenario("pending writes", ["pending_full", "127.0.0.1", port, n, 3, 65536], timeout=180)
    p.join(120)
    c.check("pending writes: queued without blocking, drained in order", p.result == gen(3, n) and
            r.get("queued") == str(n) and r.get("drained") == "True" and r.get("validAfterClose") == "True",
            "%s %s %s" % (p.error, len(p.result or b""), r))
    print("  pending writes: 8 MB queued in %s ms, drained in %s ms over %s passes" %
          (r.get("queued_ms"), r.get("drain_ms"), r.get("passes")))

    c.check("no asserts", asserts_total == 0, str(asserts_total))
    print("net: %d of %d checks pass%s" % (c.ok, c.total, "" if real_zlib else " (no zlib streams: the compressed ones skipped)"))
    if not keep:
        shutil.rmtree(os.path.join(WORK, "src"), ignore_errors=True)
    return 0 if c.ok == c.total else 1


def run(keep=False):
    """run_tests.py's entry: True when every check passes."""
    return main(["--keep"] if keep else []) == 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
