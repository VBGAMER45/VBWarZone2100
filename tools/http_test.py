"""
http_test.py - WzNet's urltransfer_c (cluster J part 2, fork LT: curl's part of src/urlrequest_curl.cpp, HTTP/1.1
over Winsock and Schannel, honouring CURLOPT_IPRESOLVE): builds build/http_test/HttpTests.exe (urltransfer_c with
WzNet's connection modules for their types, src/shared, the zlib streams of src/shared/miniz, src/nettests' log
stand-in and src/httptests' main), runs it once per scenario against a Python HTTP server on this machine and
checks what the server got and what the transfer gave:

* the address family: http://localhost with IPv4 forced arrives from 127.0.0.1, with IPv6 forced from ::1, IP_ANY
  works; an IPv4 literal with IPv6 forced does not resolve; happy eyeballs: localhost with IP_ANY when only
  127.0.0.1 listens (::1 refuses) connects over IPv4;
* the request bytes exactly as curl writes them (GET, POST and PATCH with a body, DELETE without one, a header
  replacing curl's Accept, the User-Agent);
* the bodies: Content-Length, chunked (with a trailer), read to the close, gzip, zlib deflate and raw deflate
  decoded, 204, a 100 Continue before the response;
* redirects: followed (5 hops), not followed (the 302 itself), 307 keeping the method and body, 303 dropping the
  body, too many (CURLE_TOO_MANY_REDIRECTS), to ftp (CURLE_UNSUPPORTED_PROTOCOL);
* CURLOPT_MAXFILESIZE_LARGE (a Content-Length and a chunked body past it), Retry-After (seconds and an HTTP date);
* the failures: connection refused (COULDNT_CONNECT), an empty reply, a body cut short, an unknown Content-Encoding, a TLS server that never answers (the connect timeout), a server that never
  answers the request (CURLOPT_LOW_SPEED_TIME: 60 s, with --slow).

  python -I tools\\http_test.py [--stage DIR] [--slow] [--external]

--stage: the folder holding net\\ and httptests\\ (default: src\\). --slow: the 60-second low-speed case too.
--external: a few read-only HTTPS GETs on the internet instead: the real lobby's public games list with IPv4
forced, IPv6 forced and IP_ANY (the family each used), and three that must fail certificate validation
(expired.badssl.com, wrong.host.badssl.com, self-signed.badssl.com), and a name that does not resolve
(nonexistent.invalid, the system's DNS). Without --external nothing leaves this machine.
"""
import base64
import gzip
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
import net_test  # noqa: E402

WORK = os.path.join(ROOT, "build", "http_test")
PORT = 30600            # the test server (127.0.0.1 and ::1)
PORT_V4ONLY = 30601     # 127.0.0.1 only (happy eyeballs)
PORT_REFUSED = 30602    # nothing listens
PORT_TLS_STALL = 30603  # accepts and never answers
UA = "Warzone2100/4.7.0 (WarzoneVB http_test)"


def build_exe(stage):
    src = os.path.join(WORK, "src")
    if os.path.exists(src):
        shutil.rmtree(src)
    os.makedirs(src)
    net = os.path.join(stage, "net") if os.path.exists(os.path.join(stage, "net", "urltransfer_c.bas")) else os.path.join(ROOT, "src", "net")
    files = [os.path.join(ROOT, "src", "net", n) for n in net_test.JT_MODULES]
    files.append(os.path.join(net, "urltransfer_c.bas"))
    files.append(os.path.join(ROOT, "src", "nettests", "nettests_stubs.bas"))
    tests = os.path.join(stage, "httptests")
    if not os.path.exists(tests):
        tests = os.path.join(ROOT, "src", "httptests")
    files += [os.path.join(tests, n) for n in build.sources(tests)]
    miniz = os.path.join(ROOT, "src", "shared", "miniz")
    files += [os.path.join(miniz, n) for n in build.sources(miniz)]
    files += [os.path.join(build.SHARED, n) for n in build.sources(build.SHARED)]
    lines = ["Type=Exe", build.STDOLE]
    for f in files:
        dst = os.path.join(src, os.path.basename(f))
        data = open(f, "rb").read().decode("latin-1").replace("\r\n", "\n")
        data = data.replace("Attribute VB_Exposed = True", "Attribute VB_Exposed = False")
        data = data.replace("Attribute VB_Creatable = False", "Attribute VB_Creatable = True")
        # the log stand-in writes into this exe's results
        data = data.replace("nettests_main.addResult", "httptests_main.addResult")
        open(dst, "wb").write(data.replace("\n", "\r\n").encode("latin-1"))
        name = build.vb_name(dst)
        if dst.lower().endswith(".cls"):
            lines.append("Class=%s; %s" % (name, os.path.basename(dst)))
        else:
            lines.append("Module=%s; %s" % (name, os.path.basename(dst)))
    lines.append(build.VBP_TAIL.format(name="HttpTests", ext="exe", path=".", startup="Sub Main", compat="",
                                       bounds="0", optimisation="2", threading=""))
    vbp = os.path.join(src, "HttpTests.vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    log = os.path.join(WORK, "HttpTests_build.log")
    if os.path.exists(log):
        os.remove(log)
    subprocess.run([build.VB6, "/make", vbp, "/outdir", WORK, "/out", log], cwd=src)
    text = open(log, encoding="latin-1").read().strip() if os.path.exists(log) else "VB6 stopped without a message"
    print(text)
    return "succeeded" in text


def run_transfer(url, method="GET", protocol=0, timeout=0, follow=True, maxsize=536870912, body=None, headers=(),
                 ua=UA, limit=30000):
    """HttpTests.exe on one transfer: (results dict, header lines, log lines)."""
    cf = os.path.join(WORK, "cmd.txt")
    rf = os.path.join(WORK, "results.txt")
    if os.path.exists(rf):
        os.remove(rf)
    lines = ["method\t" + method, "url\t" + url, "protocol\t%d" % protocol, "timeout\t%d" % timeout,
             "follow\t%d" % (1 if follow else 0), "maxsize\t%d" % maxsize, "ua\t" + ua, "limit\t%d" % limit]
    if body is not None:
        lines.append("body\t" + body.hex().upper())
    lines += ["header\t" + h for h in headers]
    with open(cf, "w", encoding="latin-1") as f:
        f.write("\n".join(lines) + "\n")
    subprocess.run([os.path.join(WORK, "HttpTests.exe"), cf, rf], timeout=limit / 1000 + 60)
    results, hdrs, logs = {}, [], []
    if os.path.exists(rf):
        for line in open(rf, encoding="latin-1").read().splitlines():
            parts = line.split("\t")
            if parts[0] == "result":
                results[parts[1]] = parts[2] if len(parts) > 2 else ""
            elif parts[0] == "header":
                hdrs.append(parts[1])
            elif parts[0] in ("log", "assert"):
                logs.append(line)
    if "body" in results:
        results["body_bytes"] = bytes.fromhex(results["body"])
    else:
        results["body_bytes"] = b""
    return results, hdrs, logs


class Server:
    """The test server: raw HTTP/1.1 on 127.0.0.1 and ::1 (one port), each request recorded with its peer."""

    def __init__(self):
        self.requests = []
        self.sockets = []
        self.held = []
        self.stop = False

    def listen(self, family, addr, port, handler):
        s = socket.socket(family, socket.SOCK_STREAM)
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        if family == socket.AF_INET6:
            s.setsockopt(socket.IPPROTO_IPV6, socket.IPV6_V6ONLY, 1)
        s.bind((addr, port))
        s.listen(16)
        s.settimeout(0.2)
        self.sockets.append(s)
        threading.Thread(target=self.accept_loop, args=(s, handler), daemon=True).start()

    def accept_loop(self, s, handler):
        while not self.stop:
            try:
                c, peer = s.accept()
            except socket.timeout:
                continue
            except OSError:
                return
            threading.Thread(target=handler, args=(c, peer), daemon=True).start()

    def close(self):
        self.stop = True
        for s in self.sockets + self.held:
            try:
                s.close()
            except OSError:
                pass

    @staticmethod
    def read_request(c):
        data = b""
        c.settimeout(10)
        while b"\r\n\r\n" not in data:
            chunk = c.recv(65536)
            if not chunk:
                return data
            data += chunk
        head, _, rest = data.partition(b"\r\n\r\n")
        length = 0
        for line in head.split(b"\r\n")[1:]:
            k, _, v = line.partition(b":")
            if k.strip().lower() == b"content-length":
                length = int(v.strip())
        while len(rest) < length:
            chunk = c.recv(65536)
            if not chunk:
                break
            rest += chunk
        return head + b"\r\n\r\n" + rest

    def handle(self, c, peer):
        try:
            req = self.read_request(c)
            if not req:
                c.close()
                return
            line = req.split(b"\r\n", 1)[0].decode("latin-1")
            method, path = line.split(" ")[0], line.split(" ")[1]
            self.requests.append({"peer": peer[0], "method": method, "path": path, "raw": req})
            self.respond(c, method, path, req)
        except Exception as e:  # noqa: BLE001 (the test reports what the transfer saw)
            print("  server error:", e)
            c.close()

    def respond(self, c, method, path, req):
        def send(status, body=b"", headers=(), length=True):
            out = "HTTP/1.1 %s\r\n" % status
            for h in headers:
                out += h + "\r\n"
            if length:
                out += "Content-Length: %d\r\n" % len(body)
            out = out.encode("latin-1") + b"\r\n" + body
            c.sendall(out)
            c.close()
        if path == "/echo":
            send("200 OK", req, ["Content-Type: application/octet-stream"])
        elif path == "/plain":
            send("200 OK", b"hello, plain", ["X-One: 1", "x-two: two"])
        elif path == "/chunked":
            body = b"".join(b"%x\r\n%s\r\n" % (len(p), p) for p in (b"first,", b"second," * 300, b"third")) + \
                b"0\r\nX-Trailer: done\r\n\r\n"
            c.sendall(b"HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n" + body)
            c.close()
        elif path == "/close":
            c.sendall(b"HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\n\r\n" + b"until the close " * 1000)
            c.close()
        elif path == "/gzip":
            send("200 OK", gzip.compress(b"gzipped text " * 500), ["Content-Encoding: gzip"])
        elif path == "/deflate":
            send("200 OK", zlib.compress(b"deflated text " * 500), ["Content-Encoding: deflate"])
        elif path == "/rawdeflate":
            co = zlib.compressobj(wbits=-15)
            send("200 OK", co.compress(b"raw deflate " * 500) + co.flush(), ["Content-Encoding: deflate"])
        elif path == "/badenc":
            send("200 OK", b"xxxx", ["Content-Encoding: br"])
        elif path == "/204":
            send("204 No Content", b"", length=False)
        elif path == "/continue":
            c.sendall(b"HTTP/1.1 100 Continue\r\n\r\n")
            send("200 OK", b"after continue")
        elif path.startswith("/hop/"):
            n = int(path[5:])
            if n == 0:
                send("200 OK", b"landed")
            else:
                send("302 Found", b"moved", ["Location: /hop/%d" % (n - 1)])
        elif path == "/loop":
            send("302 Found", b"", ["Location: http://localhost:%d/loop" % PORT])
        elif path == "/to307":
            send("307 Temporary Redirect", b"", ["Location: /echo"])
        elif path == "/to303":
            send("303 See Other", b"", ["Location: /echo"])
        elif path == "/toftp":
            send("302 Found", b"", ["Location: ftp://localhost/file"])
        elif path == "/big":
            send("200 OK", b"b" * 5000)
        elif path == "/bigchunked":
            c.sendall(b"HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n" + b"".join(b"3e8\r\n" + b"c" * 1000 + b"\r\n" for _ in range(5)) + b"0\r\n\r\n")
            c.close()
        elif path == "/retry":
            send("429 Too Many Requests", b"later", ["Retry-After: 7"])
        elif path == "/retrydate":
            when = time.strftime("%a, %d %b %Y %H:%M:%S GMT", time.gmtime(time.time() + 30))
            send("503 Service Unavailable", b"later", ["Retry-After: " + when])
        elif path == "/nothing":
            c.close()
        elif path == "/partial":
            c.sendall(b"HTTP/1.1 200 OK\r\nContent-Length: 100\r\n\r\n" + b"p" * 50)
            c.close()
        elif path == "/stall":
            self.held.append(c)     # never answers
        else:
            send("404 Not Found", b"no such path")


def stall_handler(server):
    def handler(c, peer):
        server.held.append(c)       # a TLS server that never answers the ClientHello
    return handler


class Checks:
    def __init__(self):
        self.ok = 0
        self.failed = []

    def check(self, name, cond, detail=""):
        if cond:
            self.ok += 1
            print("  ok   %s" % name)
        else:
            self.failed.append(name)
            print("  FAIL %s %s" % (name, detail))


def curl_request(method, target, host, extra=(), body=None, accept="*/*"):
    """The request bytes curl writes (urlrequest's options: CURLOPT_USERAGENT, CURLOPT_ACCEPT_ENCODING "")."""
    out = "%s %s HTTP/1.1\r\nHost: %s\r\nUser-Agent: %s\r\n" % (method, target, host, UA)
    if accept is not None:
        out += "Accept: %s\r\n" % accept
    out += "Accept-Encoding: deflate, gzip\r\n"
    for h in extra:
        out += h + "\r\n"
    if body is not None:
        out += "Content-Length: %d\r\n" % len(body)
        if not any(h.lower().startswith("content-type:") for h in extra):
            out += "Content-Type: application/x-www-form-urlencoded\r\n"
    return out.encode("latin-1") + b"\r\n" + (body or b"")


def local_tests(slow):
    ch = Checks()
    srv = Server()
    srv.listen(socket.AF_INET, "127.0.0.1", PORT, srv.handle)
    srv.listen(socket.AF_INET6, "::1", PORT, srv.handle)
    srv.listen(socket.AF_INET, "127.0.0.1", PORT_V4ONLY, srv.handle)
    srv.listen(socket.AF_INET, "127.0.0.1", PORT_TLS_STALL, stall_handler(srv))
    base = "http://localhost:%d" % PORT
    host = "localhost:%d" % PORT
    try:
        # the family (CURLOPT_IPRESOLVE)
        for proto, want in ((1, "127.0.0.1"), (2, "::1"), (0, None)):
            n0 = len(srv.requests)
            r, h, logs = run_transfer(base + "/plain", protocol=proto)
            peer = srv.requests[n0]["peer"] if len(srv.requests) > n0 else None
            ch.check("family %d: 200 from %s" % (proto, peer), r.get("code") == "0" and r.get("status") == "200" and
                     (want is None or peer == want) and r["body_bytes"] == b"hello, plain", str(r) + str(logs))
            ch.check("family %d: CURLINFO_PRIMARY_IP %s" % (proto, r.get("ip")), r.get("ip") == peer, r.get("ip"))
            ch.check("family %d: the connection logged with its family" % proto,
                     any("Connected to" in l and ("(IPv4)" in l if peer == "127.0.0.1" else "(IPv6)" in l) for l in logs), logs)
        r, h, logs = run_transfer("http://127.0.0.1:%d/plain" % PORT, protocol=2)
        ch.check("IPv4 literal with IPv6 forced: COULDNT_RESOLVE_HOST (6)", r.get("code") == "6", str(r))
        n0 = len(srv.requests)
        r, h, logs = run_transfer("http://localhost:%d/plain" % PORT_V4ONLY, protocol=0)
        ch.check("happy eyeballs: ::1 refuses, 127.0.0.1 answers", r.get("code") == "0" and srv.requests[n0]["peer"] == "127.0.0.1",
                 str(r) + str(logs))
        # the headers the response gave (the header callback's lines)
        r, h, logs = run_transfer(base + "/plain")
        ch.check("response header lines", h[:1] == ["HTTP/1.1 200 OK"] and "X-One: 1" in h and "x-two: two" in h, h)

        # the request bytes
        cases = [
            ("GET", "/echo", None, [], "*/*"),
            ("POST", "/echo", b'{"a":1,"text":"\xc3\xa9t\xc3\xa9"}', ["X-Test: one", "Content-Type: application/json"], "*/*"),
            ("PATCH", "/echo", b'{"patch":true}', ["Content-Type: application/json"], "*/*"),
            ("DELETE", "/echo", None, ["X-Host-Token: abc"], "*/*"),
            ("GET", "/echo", None, ["Accept: application/jsonl", "Origin: wz2100://wz2100.net"], None),
        ]
        for method, path, body, extra, accept in cases:
            n0 = len(srv.requests)
            r, h, logs = run_transfer(base + path, method=method, body=body, headers=extra)
            got = srv.requests[n0]["raw"] if len(srv.requests) > n0 else b""
            want = curl_request(method, path, host, extra, body, accept)
            ch.check("%s request bytes as curl's" % method, got == want, "\n got %r\nwant %r" % (got, want))
            ch.check("%s echoed body" % method, r["body_bytes"] == got, str(r)[:200])

        # bodies
        r, h, logs = run_transfer(base + "/chunked")
        ch.check("chunked body", r["body_bytes"] == b"first," + b"second," * 300 + b"third", r.get("bodylen"))
        ch.check("chunked trailer in the header lines", "X-Trailer: done" in h, h)
        r, h, logs = run_transfer(base + "/close")
        ch.check("body read to the close", r.get("code") == "0" and r["body_bytes"] == b"until the close " * 1000, r.get("bodylen"))
        for path, want in (("/gzip", b"gzipped text " * 500), ("/deflate", b"deflated text " * 500), ("/rawdeflate", b"raw deflate " * 500)):
            r, h, logs = run_transfer(base + path)
            ch.check("%s decoded" % path, r.get("code") == "0" and r["body_bytes"] == want, str(r)[:200] + str(logs))
        r, h, logs = run_transfer(base + "/204")
        ch.check("204: no body", r.get("code") == "0" and r.get("status") == "204" and r.get("bodylen") == "0", str(r))
        r, h, logs = run_transfer(base + "/continue")
        ch.check("100 Continue skipped", r.get("status") == "200" and r["body_bytes"] == b"after continue" and
                 "HTTP/1.1 100 Continue" in h, str(r) + str(h))

        # redirects
        r, h, logs = run_transfer(base + "/hop/5")
        ch.check("5 redirects followed", r.get("code") == "0" and r["body_bytes"] == b"landed" and
                 sum(1 for l in h if l.startswith("HTTP/1.1 302")) == 5, str(r) + str(h))
        r, h, logs = run_transfer(base + "/hop/5", follow=False)
        ch.check("not followed: the 302 itself", r.get("status") == "302" and r["body_bytes"] == b"moved", str(r))
        r, h, logs = run_transfer(base + "/loop")
        ch.check("too many redirects: 47", r.get("code") == "47", str(r))
        n0 = len(srv.requests)
        r, h, logs = run_transfer(base + "/to307", method="POST", body=b"keep me", headers=["Content-Type: text/plain"])
        ch.check("307 keeps POST and the body", len(srv.requests) == n0 + 2 and srv.requests[n0 + 1]["method"] == "POST" and
                 srv.requests[n0 + 1]["raw"].endswith(b"\r\n\r\nkeep me"), [q["raw"] for q in srv.requests[n0:]])
        n0 = len(srv.requests)
        r, h, logs = run_transfer(base + "/to303", method="POST", body=b"drop me", headers=["Content-Type: text/plain"])
        second = srv.requests[n0 + 1]["raw"] if len(srv.requests) > n0 + 1 else b""
        ch.check("303 drops the body (curl 8: the custom method stays)", second.startswith(b"POST /echo") and
                 b"Content-Length" not in second and second.endswith(b"\r\n\r\n"), second)
        r, h, logs = run_transfer(base + "/toftp")
        ch.check("redirect to ftp: 1", r.get("code") == "1", str(r))

        # limits and Retry-After
        r, h, logs = run_transfer(base + "/big", maxsize=1000)
        ch.check("Content-Length past CURLOPT_MAXFILESIZE: 63", r.get("code") == "63", str(r))
        r, h, logs = run_transfer(base + "/bigchunked", maxsize=1000)
        ch.check("chunked body growing past it: 63", r.get("code") == "63", str(r))
        r, h, logs = run_transfer(base + "/big", maxsize=5000)
        ch.check("a body of exactly CURLOPT_MAXFILESIZE: kept", r.get("code") == "0" and r.get("bodylen") == "5000", str(r))
        r, h, logs = run_transfer(base + "/retry")
        ch.check("Retry-After seconds: 7", r.get("status") == "429" and r.get("retryafter") == "7", str(r))
        r, h, logs = run_transfer(base + "/retrydate")
        ra = float(r.get("retryafter") or 0)
        ch.check("Retry-After date: about 30 s (%s)" % ra, 25 <= ra <= 31, str(r))

        # failures
        r, h, logs = run_transfer("http://127.0.0.1:%d/" % PORT_REFUSED, protocol=1)
        ch.check("refused: COULDNT_CONNECT (7)", r.get("code") == "7", str(r))
        r, h, logs = run_transfer(base + "/nothing")
        ch.check("empty reply: GOT_NOTHING (52)", r.get("code") == "52", str(r))
        r, h, logs = run_transfer(base + "/partial")
        ch.check("body cut short: PARTIAL_FILE (18)", r.get("code") == "18", str(r))
        r, h, logs = run_transfer(base + "/badenc")
        ch.check("unknown Content-Encoding: 61", r.get("code") == "61", str(r))
        t0 = time.time()
        r, h, logs = run_transfer("https://127.0.0.1:%d/" % PORT_TLS_STALL, timeout=2000)
        dt = time.time() - t0
        ch.check("TLS server that never answers: OPERATION_TIMEDOUT (28) after the 2 s connect timeout (%.1f s)" % dt,
                 r.get("code") == "28" and 1.5 < float(r.get("elapsed", 0)) / 1000 < 4, str(r) + str(logs))
        if slow:
            r, h, logs = run_transfer(base + "/stall", limit=90000)
            el = float(r.get("elapsed", 0)) / 1000
            ch.check("no answer: low-speed limit, 28 after about 60 s (%.0f s)" % el, r.get("code") == "28" and 58 <= el <= 70,
                     str(r) + str(logs))
    finally:
        srv.close()
    return ch


def external_tests():
    """Read-only HTTPS GETs: the real lobby's public games list (each family) and badssl.com's broken certificates."""
    ch = Checks()
    lobby = "https://wzlobby.wz2100.net/api/v1/games"
    hdrs = ["Origin: wz2100://wz2100.net", "Referer: https://github.com/Warzone2100/warzone2100?wz_internal=list",
            "Accept: application/jsonl"]
    for proto, fam in ((1, "(IPv4)"), (2, "(IPv6)"), (0, None)):
        r, h, logs = run_transfer(lobby, protocol=proto, follow=False, headers=hdrs, timeout=10000, maxsize=20 * 1024 * 1024,
                                  ua="Warzone2100/4.7.0 (+https://github.com/Warzone2100/warzone2100)")
        conn = [l for l in logs if "Connected to" in l]
        tls = [l for l in logs if "schannel:" in l]
        print("    %s | %s | status %s, %s bytes" % (conn[-1].split("\t")[-1] if conn else "-", tls[-1].split("\t")[-1] if tls else "-",
                                                  r.get("status"), r.get("bodylen")))
        # the lobby's answer over TLS on the forced family: its list (200), or Cloudflare's challenge page (403 with
        # Cf-Mitigated: challenge), which curl.exe and WinHttpRequest got too on 2026-11-01 (not this transport's doing)
        challenged = r.get("status") == "403" and any(l.lower().startswith("cf-mitigated: challenge") for l in h)
        ok = r.get("code") == "0" and (r.get("status") == "200" or challenged) and int(r.get("bodylen", 0)) > 0
        if fam:
            ok = ok and any(fam in l for l in conn)
        ch.check("real lobby games list, protocol %d%s" % (proto, " (Cloudflare challenged it)" if challenged else ""), ok,
                 str({k: v for k, v in r.items() if k != "body" and k != "body_bytes"}) + str(logs))
        if r.get("status") == "200":
            first = r["body_bytes"].split(b"\n", 1)[0]
            print("    first line: %s" % first[:160])
    r, h, logs = run_transfer("http://nonexistent.invalid/")
    ch.check("unknown name (the system's DNS): COULDNT_RESOLVE_HOST (6)", r.get("code") == "6", str(r))
    for host, want in (("expired.badssl.com", "35"), ("wrong.host.badssl.com", "60"), ("self-signed.badssl.com", "35")):
        r, h, logs = run_transfer("https://%s/" % host, timeout=10000)
        detail = [l.split("\t")[-1] for l in logs if "(VB)" in l]
        print("    %s: code %s %s" % (host, r.get("code"), detail[-1] if detail else ""))
        ch.check("%s refused (code %s)" % (host, want), r.get("code") == want and r.get("status") == "0", str(r))
    return ch


def main(argv):
    stage = os.path.join(ROOT, "src")
    if "--stage" in argv:
        stage = argv[argv.index("--stage") + 1]
    os.makedirs(WORK, exist_ok=True)
    if not build_exe(stage):
        return 1
    ch = external_tests() if "--external" in argv else local_tests("--slow" in argv)
    print("%d ok, %d failed%s" % (ch.ok, len(ch.failed), (": " + ", ".join(ch.failed)) if ch.failed else ""))
    return 1 if ch.failed else 0


def run(keep=False):
    """run_tests.py's entry (the local cases only: nothing leaves this machine): True when every check passes."""
    return main([]) == 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
