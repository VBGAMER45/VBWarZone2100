"""
crypto_test.py - cluster J's test of src/shared/sodium: lib/framework/crc.cpp's sha256Sum, EcKey, SessionKeys,
base64 and genSecRandomBytes, and the libsodium primitives under them (SHA-256, SHA-512, BLAKE2b, ChaCha20,
HChaCha20, Poly1305, XChaCha20-Poly1305 IETF, X25519, Ed25519, crypto_kx, sodium_pad), against libsodium itself
through PyNaCl (nacl.bindings), hashlib and the `cryptography` package:

* builds CryptoTests.exe (src/cryptotests + src/shared/sodium + src/shared's math_vb, crc_c and utf8_vb; no DLL;
  compiled as the exe is: no optimisation, bounds checks) and runs it on a command file;
* the hashes on every length 0 .. 300 and the block boundaries, whole and fed in chunks; BLAKE2b at every
  digest length; the stream cipher, Poly1305 and the AEAD on random lengths (and RFC/draft test vectors);
  the AEAD's tampered ciphertexts, tags, nonces and associated data refused;
* X25519 on random scalars and points and on the small-order points libsodium refuses; Ed25519 seed keypairs,
  signatures (random messages, empty included) and their verification, the verification of tampered messages
  and signatures and of the non-canonical and small-order cases libsodium refuses; is_valid_point on public keys,
  small-order and non-canonical encodings and off-curve values; the keys' conversion to Curve25519; crypto_kx
  both ways; sodium_pad / sodium_unpad;
* crc.cpp: sha256Sum; EcKey from public and private bytes (valid and invalid, with crc.cpp's log lines),
  generate, sign / verify, publicHashString (truncated too), publicKeyHexString; SessionKeys between two keys
  (client and server chosen by the hex strings, the player index when equal), its keys against a Python version
  of crc.cpp's SessionKeys over PyNaCl, messages the port encrypts decrypted by PyNaCl and the reverse, the
  padding to 64 bytes, forged messages refused; base64 against Python's (and crc.cpp's own decoder rules);
* a speed line (per operation, in ms).

  python -I tools\\crypto_test.py [--keep] [--stage DIR] [--out DIR]   (run_tests.py calls run(keep))
  --stage: the sources from DIR (a mirror of src/, e.g. a fork's staging folder) instead of src/;
  --out: CryptoTests.exe into DIR instead of bin/.
"""
import base64
import hashlib
import math
import os
import random
import shutil
import subprocess
import sys
import time

ROOT = r"C:\Users\Owner\desktop\websites\warzone"
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

import nacl.bindings as nb  # noqa: E402
import nacl.exceptions  # noqa: E402
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms  # noqa: E402
from cryptography.hazmat.primitives import poly1305  # noqa: E402

WORK = os.path.join(ROOT, "build", "crypto_test")
L = 2 ** 252 + 27742317777372353535851937790883648493
P = 2 ** 255 - 19


def build_exe(src, out):
    folder = os.path.join(src, "cryptotests")
    files = [os.path.join(folder, n) for n in build.sources(folder)]
    files += [os.path.join(src, "shared", "sodium", n) for n in build.sources(os.path.join(src, "shared", "sodium"))]
    for n in ("math_vb.bas", "crc_c.bas", "utf8_vb.bas"):
        files.append(os.path.join(build.SHARED, n))
    lines = ["Type=Exe", build.STDOLE]
    for p in files:
        build.normalise(p)
        lines.append("Module=%s; %s" % (build.vb_name(p), p))
    # sodium_25519_c records the largest magnitude its Doubles reach (limbmax): the exactness check
    lines.append('CondComp="SODIUM_LIMB_CHECK = 1"')
    lines.append(build.VBP_TAIL.format(name="CryptoTests", ext="exe", path=out, startup="Sub Main",
                                       compat="", bounds="0", optimisation="2", threading=""))
    os.makedirs(WORK, exist_ok=True)
    vbp = os.path.join(WORK, "CryptoTests.vbp")
    with open(vbp, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(lines))
    saved = build.BIN
    build.BIN = out
    try:
        return build.make("CryptoTests", WORK, vbp)
    finally:
        build.BIN = saved


def run_exe(exe, commands):
    """Runs the commands; returns, per command, (its result line, the LOG / ASSERT lines before it)."""
    cmd = os.path.join(WORK, "commands.txt")
    res = os.path.join(WORK, "results.txt")
    with open(cmd, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(commands) + "\n")
    if os.path.exists(res):
        os.remove(res)
    t0 = time.time()
    subprocess.run([exe, cmd, res], timeout=3600)
    dt = time.time() - t0
    out = []
    logs = []
    for ln in open(res, encoding="utf-8").read().split("\n")[:-1]:     # (an empty result is a line too)
        if ln.startswith("LOG\t") or ln.startswith("ASSERT\t"):
            logs.append(ln)
        else:
            out.append((ln, logs))
            logs = []
    if len(out) != len(commands):
        raise RuntimeError("CryptoTests.exe answered %d of %d commands" % (len(out), len(commands)))
    return out, dt


def h(b):
    return b.hex() if b else "-"


# ---------------------------------------------------------------------------------------------------------
# references

def chacha20_ietf(key, nonce12, ic, m):
    c = Cipher(algorithms.ChaCha20(key, ic.to_bytes(4, "little") + nonce12), mode=None).encryptor()
    return c.update(m)


def poly(key, m):
    p = poly1305.Poly1305(key)
    p.update(m)
    return p.finalize()


def clamp(sk):
    b = bytearray(sk)
    b[0] &= 248
    b[31] &= 127
    b[31] |= 64
    return bytes(b)


def x25519_ref(n, p):
    try:
        return nb.crypto_scalarmult(n, p)
    except Exception:
        return None


def ed_verify_ref(sig, m, pk):
    try:
        nb.crypto_sign_open(sig + m, pk)
        return 0
    except Exception:
        return -1


def pk2c_ref(pk):
    try:
        return nb.crypto_sign_ed25519_pk_to_curve25519(pk)
    except Exception:
        return None


def enc_point(y, sign=0):
    return (y | (sign << 255)).to_bytes(32, "little")


class PySessionKeys:
    """crc.cpp's SessionKeys over PyNaCl."""

    def __init__(self, me_sk, me_idx, other_pk, other_idx):
        me_pk = me_sk[32:]
        a, b = me_pk.hex(), other_pk.hex()
        client = (a < b) if a != b else (me_idx < other_idx)
        me_cpk = nb.crypto_sign_ed25519_pk_to_curve25519(me_pk)
        me_csk = nb.crypto_sign_ed25519_sk_to_curve25519(me_sk)
        other_cpk = nb.crypto_sign_ed25519_pk_to_curve25519(other_pk)
        if client:
            self.rx, self.tx = nb.crypto_kx_client_session_keys(me_cpk, me_csk, other_cpk)
        else:
            self.rx, self.tx = nb.crypto_kx_server_session_keys(me_cpk, me_csk, other_cpk)

    def encrypt(self, m):
        padded = nb.sodium_pad(m, 64)
        nonce = os.urandom(24)
        return nonce + nb.crypto_aead_xchacha20poly1305_ietf_encrypt(padded, None, nonce, self.tx)

    def decrypt(self, data):
        try:
            padded = nb.crypto_aead_xchacha20poly1305_ietf_decrypt(data[24:], None, data[:24], self.rx)
        except Exception:
            return None
        return nb.sodium_unpad(padded, 64)


def cpp_base64_decode(s):
    """crc.cpp's base64Decode, literally (whole quads only; '=' trims at the last quad's end)."""
    b = s.encode("utf-8")
    out = bytearray(len(b) // 4 * 3)
    for n in range(len(b) // 4):
        block = 0
        for i in range(4):
            ch = b[i + n * 4]
            if 65 <= ch <= 90:
                v = ch - 65
            elif 97 <= ch <= 122:
                v = ch - 97 + 26
            elif 48 <= ch <= 57:
                v = ch - 48 + 52
            elif ch == 43:
                v = 62
            elif ch == 47:
                v = 63
            else:
                v = 0
            block |= v << (6 * (3 - i))
        out[n * 3:n * 3 + 3] = bytes([(block >> 16) & 255, (block >> 8) & 255, block & 255])
    if len(b) >= 4:
        if b[len(b) // 4 * 4 - 2] == 61:
            out = out[:-2]
        elif b[len(b) // 4 * 4 - 1] == 61:
            out = out[:-1]
    return bytes(out)


# ---------------------------------------------------------------------------------------------------------

def run(keep=False, stage=None, out=None):
    rng = random.Random(4700)
    src = stage or build.SRC
    out = out or build.BIN
    if not build_exe(src, out):
        print("crypto: CryptoTests.exe did not build")
        return False
    exe = os.path.join(out, "CryptoTests.exe")
    cmds = []
    checks = []          # (label, expected result line or a function of the line and its logs)

    def add(cmd, label, expect):
        cmds.append(cmd)
        checks.append((label, expect))

    def rb(n):
        return bytes(rng.getrandbits(8) for _ in range(n))

    # hashes
    for n in list(range(0, 301)) + [511, 512, 513, 1000, 4096 + 7]:
        m = rb(n)
        add("sha256\t" + h(m), "sha256 %d" % n, hashlib.sha256(m).hexdigest())
        add("sha512\t" + h(m), "sha512 %d" % n, hashlib.sha512(m).hexdigest())
    for n, chunk in ((1000, 1), (1000, 7), (1000, 63), (1000, 64), (1000, 65), (1000, 127), (1000, 129), (300, 300)):
        m = rb(n)
        add("sha256chunks\t%s\t%d" % (h(m), chunk), "sha256 chunks %d/%d" % (n, chunk), hashlib.sha256(m).hexdigest())
        add("sha512chunks\t%s\t%d" % (h(m), chunk), "sha512 chunks %d/%d" % (n, chunk), hashlib.sha512(m).hexdigest())
    for n in (0, 1, 31, 32, 63, 64, 95, 96, 127, 128, 129, 255, 256, 257, 1000):
        m = rb(n)
        for outlen in (1, 16, 32, 63, 64):
            add("blake2b\t%d\t%s" % (outlen, h(m)), "blake2b %d/%d" % (n, outlen),
                hashlib.blake2b(m, digest_size=outlen).hexdigest())
    # HChaCha20: draft-irtf-cfrg-xchacha-03, 2.2.1
    add("hchacha20\t" + bytes(range(32)).hex() + "\t000000090000004a0000000031415927", "hchacha20 draft vector",
        "82413b4227b27bfed30e42508a877d73a0f9e4d58a74a853c12ec41326d3ecdc")
    for n in (0, 1, 63, 64, 65, 200, 1000):
        k, nonce, m = rb(32), rb(12), rb(n)
        ic = rng.choice([0, 1, 7, 0xFFFFFFFE])
        if n > 64:
            ic = rng.choice([0, 1, 7])
        add("chacha20\t%s\t%s\t%d\t%s" % (h(k), h(nonce), ic, h(m)), "chacha20 %d ic %d" % (n, ic),
            h(chacha20_ietf(k, nonce, ic, m)))
    for n in (0, 1, 15, 16, 17, 31, 32, 33, 64, 100, 1000):
        k, m = rb(32), rb(n)
        add("poly1305\t%s\t%s" % (h(k), h(m)), "poly1305 %d" % n, h(poly(k, m)))
    k = bytes.fromhex("85d6be7857556d337f4452fe42d506a80103808afb0db2fd4abff6af4149f51b")      # RFC 8439 2.5.2
    add("poly1305\t%s\t%s" % (h(k), h(b"Cryptographic Forum Research Group")), "poly1305 RFC 8439",
        "a8061dc1305136c6c22b8baf0c0127a9")
    k, m = bytes(16) + rb(16), rb(100)                 # r = 0: the tag is s
    add("poly1305\t%s\t%s" % (h(k), h(m)), "poly1305 r = 0", h(poly(k, m)))
    for n, adn in ((0, 0), (1, 0), (0, 5), (63, 0), (64, 12), (65, 17), (200, 3), (1000, 0)):
        k, nonce, m, ad = rb(32), rb(24), rb(n), rb(adn)
        c = nb.crypto_aead_xchacha20poly1305_ietf_encrypt(m, ad or None, nonce, k)
        add("aead_enc\t%s\t%s\t%s\t%s" % (h(k), h(nonce), h(ad), h(m)), "aead enc %d/%d" % (n, adn), h(c))
        add("aead_dec\t%s\t%s\t%s\t%s" % (h(k), h(nonce), h(ad), h(c)), "aead dec %d/%d" % (n, adn), h(m))
        bad = bytearray(c)
        bad[rng.randrange(len(bad))] ^= 1 << rng.randrange(8)
        add("aead_dec\t%s\t%s\t%s\t%s" % (h(k), h(nonce), h(ad), h(bytes(bad))), "aead tampered %d" % n, "FAIL")
        add("aead_dec\t%s\t%s\t%s\t%s" % (h(k), h(rb(24)), h(ad), h(c)), "aead other nonce %d" % n, "FAIL")
        add("aead_dec\t%s\t%s\t%s\t%s" % (h(k), h(nonce), h(ad + b"x"), h(c)), "aead other ad %d" % n, "FAIL")
    add("aead_dec\t%s\t%s\t-\t%s" % (h(rb(32)), h(rb(24)), h(rb(15))), "aead too short", "FAIL")

    # X25519
    for i in range(12):
        n, p = rb(32), rb(32)
        if i < 3:
            p = nb.crypto_scalarmult_base(rb(32))
        r = x25519_ref(n, p)
        add("x25519\t%s\t%s" % (h(n), h(p)), "x25519 %d" % i, h(r) if r else "FAIL")
    small = ["00" * 32, "01" + "00" * 31, "e0eb7a7c3b41b8ae1656e3faf19fc46ada098deb9c32b1fd866205165f49b800",
             "5f9c95bca3508c24b1d0b1559c83ef5b04445cc4581c8e86d8224eddd09f1157", "ec" + "ff" * 30 + "7f",
             "ed" + "ff" * 30 + "7f", "ee" + "ff" * 30 + "7f"]
    for s in small:
        for top in (0, 0x80):
            p = bytearray(bytes.fromhex(s))
            p[31] |= top
            p = bytes(p)
            r = x25519_ref(rb(32), p)
            add("x25519\t%s\t%s" % (h(rb(32)), h(p)), "x25519 small order %s..%02x" % (s[:8], p[31]),
                "FAIL" if r is None else None)

    # Ed25519
    keys = []
    for i in range(6):
        seed = rb(32)
        pk, sk = nb.crypto_sign_seed_keypair(seed)
        keys.append((pk, sk))
        add("ed_seedkp\t" + h(seed), "ed seed keypair %d" % i, h(pk) + "\t" + h(sk))
    for i, (pk, sk) in enumerate(keys):
        for n in (0, 1, 32, 100, 333):
            m = rb(n)
            sig = nb.crypto_sign(m, sk)[:64]
            add("ed_sign\t%s\t%s" % (h(sk), h(m)), "ed sign key %d len %d" % (i, n), h(sig))
            add("ed_verify\t%s\t%s\t%s" % (h(sig), h(m), h(pk)), "ed verify key %d len %d" % (i, n), "0")
            m2 = bytearray(m + b"!")
            add("ed_verify\t%s\t%s\t%s" % (h(sig), h(bytes(m2)), h(pk)), "ed verify other message", "-1")
            bad = bytearray(sig)
            bad[rng.randrange(64)] ^= 1 << rng.randrange(8)
            exp = ed_verify_ref(bytes(bad), m, pk)
            add("ed_verify\t%s\t%s\t%s" % (h(bytes(bad)), h(m), h(pk)), "ed verify tampered sig", str(exp))
            other = keys[(i + 1) % len(keys)][0]
            add("ed_verify\t%s\t%s\t%s" % (h(sig), h(m), h(other)), "ed verify other key", "-1")
    # the edge cases: S + L (non-canonical S), the identity as R, small-order and non-canonical public keys
    pk, sk = keys[0]
    m = b"edge"
    sig = nb.crypto_sign(m, sk)[:64]
    s = int.from_bytes(sig[32:], "little") + L
    if s < 2 ** 256:
        sig2 = sig[:32] + s.to_bytes(32, "little")
        add("ed_verify\t%s\t%s\t%s" % (h(sig2), h(m), h(pk)), "ed verify S + L", str(ed_verify_ref(sig2, m, pk)))
    ident = enc_point(1)
    sig3 = ident + sig[32:]
    add("ed_verify\t%s\t%s\t%s" % (h(sig3), h(m), h(pk)), "ed verify R = identity", str(ed_verify_ref(sig3, m, pk)))
    for y in (1, P - 1, 0, P, P + 1, P + 3):
        if y >= 2 ** 255:
            continue
        for sign in (0, 1):
            e = enc_point(y, sign)
            add("ed_verify\t%s\t%s\t%s" % (h(sig), h(m), h(e)), "ed verify pk y=%d sign %d" % (y if y < 100 else y - P, sign),
                str(ed_verify_ref(sig, m, e)))
            add("ed_valid\t" + h(e), "is_valid_point y=%d sign %d" % (y if y < 100 else y - P, sign),
                str(int(nb.crypto_core_ed25519_is_valid_point(e))))
            r = pk2c_ref(e)
            add("ed_pk2c\t" + h(e), "pk to curve y=%d sign %d" % (y if y < 100 else y - P, sign), h(r) if r else "FAIL")
    for i in range(20):
        e = rb(32)
        add("ed_valid\t" + h(e), "is_valid_point random %d" % i, str(int(nb.crypto_core_ed25519_is_valid_point(e))))
        r = pk2c_ref(e)
        add("ed_pk2c\t" + h(e), "pk to curve random %d" % i, h(r) if r else "FAIL")
    # a point of the curve outside the main subgroup: pk + a point of order 8
    for i, (pk, sk) in enumerate(keys):
        add("ed_valid\t" + h(pk), "is_valid_point key %d" % i, "1")
        add("ed_pk2c\t" + h(pk), "pk to curve key %d" % i, h(nb.crypto_sign_ed25519_pk_to_curve25519(pk)))
        add("ed_sk2c\t" + h(sk), "sk to curve key %d" % i, h(nb.crypto_sign_ed25519_sk_to_curve25519(sk)))
        try:
            mixed = nb.crypto_core_ed25519_add(pk, bytes.fromhex("26e8958fc2b227b045c3f489f2ef98f0d5dfac05d3c63339b13802886d53fc05"))
            add("ed_valid\t" + h(mixed), "is_valid_point key %d + order 8" % i, str(int(nb.crypto_core_ed25519_is_valid_point(mixed))))
            r = pk2c_ref(mixed)
            add("ed_pk2c\t" + h(mixed), "pk to curve key %d + order 8" % i, h(r) if r else "FAIL")
        except Exception:
            pass
    # crypto_kx
    for i in range(6):
        csk, ssk = rb(32), rb(32)
        cpk, spk = nb.crypto_scalarmult_base(csk), nb.crypto_scalarmult_base(ssk)
        rx, tx = nb.crypto_kx_client_session_keys(cpk, csk, spk)
        add("kx_client\t%s\t%s\t%s" % (h(cpk), h(csk), h(spk)), "kx client %d" % i, h(rx) + "\t" + h(tx))
        rx, tx = nb.crypto_kx_server_session_keys(spk, ssk, cpk)
        add("kx_server\t%s\t%s\t%s" % (h(spk), h(ssk), h(cpk)), "kx server %d" % i, h(rx) + "\t" + h(tx))
    add("kx_client\t%s\t%s\t%s" % (h(rb(32)), h(rb(32)), "00" * 32), "kx small-order server key", "FAIL")
    # sodium_pad / unpad
    for n in (0, 1, 63, 64, 65, 127, 128, 200):
        m = rb(n)
        maxlen = n + (64 - n % 64)
        buf = m + bytes(maxlen - n)
        padded = nb.sodium_pad(m, 64)
        add("pad\t%s\t%d\t64\t%d" % (h(buf), n, maxlen), "sodium_pad %d" % n, "0\t%d\t%s" % (len(padded), h(padded)))
        add("unpad\t%s\t64" % h(padded), "sodium_unpad %d" % n, "0\t%d" % n)
    add("unpad\t%s\t64" % h(rb(63) + b"\x01"), "sodium_unpad no marker", lambda ln, lg: ln.startswith("-1\t"))
    add("unpad\t%s\t64" % h(bytes(64)), "sodium_unpad all zero", lambda ln, lg: ln.startswith("-1\t"))
    add("pad\t%s\t64\t64\t64" % h(rb(64)), "sodium_pad no room", lambda ln, lg: ln.startswith("-1\t"))

    # crc.cpp
    for n in (0, 1, 64, 1000):
        m = rb(n)
        add("sha256sum\t" + h(m), "sha256Sum %d" % n, hashlib.sha256(m).hexdigest())
    add("eckey_empty", "EcKey empty",
        lambda ln, lg: ln == "True\tFalse\t\t\t0\tFalse" and any("\t258\t" in x and "No key" in x for x in lg)
        and any("\t289\t" in x for x in lg) and sum("\t321\t" in x for x in lg) == 2)
    add("eckey_gen", "EcKey generate",
        lambda ln, lg: len(ln.split("\t")[0]) == 128 and ln.split("\t")[0][64:] == ln.split("\t")[1]
        and nb.crypto_sign_seed_keypair(bytes.fromhex(ln.split("\t")[0][:64]))[0].hex() == ln.split("\t")[1]
        and ln.endswith("True"))
    for i, (pk, sk) in enumerate(keys[:3]):
        hs = hashlib.sha256(pk).hexdigest()
        add("eckey_pub\t" + h(pk), "EcKey public %d" % i,
            lambda ln, lg, pk=pk, hs=hs: ln.split("\t") == ["True", pk.hex(), "0", hs, hs[:10], pk.hex(), "False"]
            and any("\t335\t" in x for x in lg))
        add("eckey_priv\t" + h(sk), "EcKey private %d" % i, "True\t%s\t%s\tTrue" % (h(pk), h(sk)))
        m = rb(50)
        sig = nb.crypto_sign(m, sk)[:64]
        add("eckey_sign\t%s\t%s" % (h(sk), h(m)), "EcKey sign %d" % i, h(sig))
        add("eckey_verify\t%s\t%s\t%s" % (h(pk), h(sig), h(m)), "EcKey verify %d" % i, "True")
        add("eckey_verify\t%s\t%s\t%s" % (h(pk), h(sig[:63]), h(m)), "EcKey verify short sig %d" % i,
            lambda ln, lg: ln == "False" and any("\t295\t" in x and "wrong size" in x and x.endswith("\t63\t64") for x in lg))
    add("eckey_pub\t" + h(rb(31)), "EcKey public wrong size",
        lambda ln, lg: ln.startswith("False\t-\t0\t\t\t\tFalse") and any("\t356\t" in x for x in lg))
    add("eckey_pub\t" + h(enc_point(1)), "EcKey public identity",
        lambda ln, lg: ln.startswith("False\t") and any("\t362\t" in x for x in lg))
    add("eckey_priv\t" + h(rb(63)), "EcKey private wrong size",
        lambda ln, lg: ln.startswith("False\t") and any("\t374\t" in x for x in lg))
    badsk = keys[0][1][:32] + enc_point(1)
    add("eckey_priv\t" + h(badsk), "EcKey private bad keypair",
        lambda ln, lg: ln.startswith("False\t") and any("\t389\t" in x for x in lg))
    # SessionKeys
    for i in range(4):
        (apk, ask), (bpk, bsk) = keys[i], keys[i + 1]
        mine = PySessionKeys(ask, 1, bpk, 2)
        other = PySessionKeys(bsk, 2, apk, 1)
        add("sk_init\t%s\t1\t%s\t2" % (h(ask), h(bpk)), "SessionKeys %d" % i, "1\t%s\t%s" % (h(mine.rx), h(mine.tx)))
        for n in (0, 1, 63, 64, 65, 300):
            m = rb(n)
            add("sk_enc\t%s\t1\t%s\t2\t%s" % (h(ask), h(bpk), h(m)), "SessionKeys %d encrypt %d" % (i, n),
                lambda ln, lg, other=other, m=m: ln != "-" and len(bytes.fromhex(ln)) == 24 + (len(m) // 64 + 1) * 64 + 16
                and other.decrypt(bytes.fromhex(ln)) == m)
            data = other.encrypt(m)
            add("sk_dec\t%s\t1\t%s\t2\t%s" % (h(ask), h(bpk), h(data)), "SessionKeys %d decrypt %d" % (i, n), "1\t" + h(m))
        forged = bytearray(other.encrypt(b"hello"))
        forged[30] ^= 4
        add("sk_dec\t%s\t1\t%s\t2\t%s" % (h(ask), h(bpk), h(bytes(forged))), "SessionKeys %d forged" % i,
            lambda ln, lg: ln == "0" and any("crc.cpp:682" in x for x in lg))
        add("sk_dec\t%s\t1\t%s\t2\t%s" % (h(ask), h(bpk), h(rb(30))), "SessionKeys %d too short" % i,
            lambda ln, lg: ln == "0" and any("crc.cpp:670" in x for x in lg))
    pk, sk = keys[0]
    same_lo = PySessionKeys(sk, 1, pk, 3)
    add("sk_init\t%s\t1\t%s\t3" % (h(sk), h(pk)), "SessionKeys same key, index decides", "1\t%s\t%s" % (h(same_lo.rx), h(same_lo.tx)))
    same_hi = PySessionKeys(sk, 3, pk, 1)
    add("sk_init\t%s\t3\t%s\t1" % (h(sk), h(pk)), "SessionKeys same key, other index", "1\t%s\t%s" % (h(same_hi.rx), h(same_hi.tx)))
    add("sk_init\t-\t1\t%s\t2" % h(pk), "SessionKeys me without a private key",
        lambda ln, lg: ln == '0\tInvalid "me" key')
    add("sk_init\t%s\t1\t%s\t2" % (h(sk), h(enc_point(1))), "SessionKeys other the identity",
        lambda ln, lg: ln == '0\tInvalid "other" public key')
    # base64
    for n in (0, 1, 2, 3, 4, 5, 32, 64, 100):
        m = rb(n)
        enc = base64.b64encode(m).decode()
        add("b64enc\t" + h(m), "base64Encode %d" % n, enc)
        add("b64dec\t" + (enc or "-"), "base64Decode %d" % n, h(m))
    for s in ("QUJD", "QUI=", "QQ==", "QUJDRA", "QU", "QUJD=x", "####"):
        add("b64dec\t" + s, "base64Decode %s" % s, h(cpp_base64_decode(s)))
    add("b64url\ta+b/c+/=", "b64Tob64UrlSafe", "a-b_c-_=")
    add("b64std\ta-b_c-_=", "b64UrlSafeTob64", "a+b/c+/=")
    add("random\t64", "genSecRandomBytes", lambda ln, lg: len(ln) == 128 and len(set(ln)) > 8)
    add("random\t64", "genSecRandomBytes again", lambda ln, lg, cmds=cmds: len(ln) == 128)
    add("speed\t5", "speed", lambda ln, lg: True)
    add("limbmax", "limbmax", lambda ln, lg: 0 < float(ln) < 2.0 ** 53)

    results, dt = run_exe(exe, cmds)
    fails = 0
    speed = ""
    limb = 0.0
    randoms = []
    for (ln, logs), (label, expect), cmd in zip(results, checks, cmds):
        if label == "speed":
            speed = ln
            continue
        if label == "limbmax":
            limb = float(ln) if ln else 0.0
        if label.startswith("genSecRandomBytes"):
            randoms.append(ln)
        if expect is None:
            ok = True
        elif callable(expect):
            ok = expect(ln, logs)
        else:
            ok = (ln == expect)
        if not ok:
            fails += 1
            if fails <= 25:
                print("  FAIL %s: got %s, expected %s%s" % (label, ln[:200], expect if not callable(expect) else "(check)",
                                                        ("; logs " + " | ".join(logs))[:300] if logs else ""))
    if len(randoms) == 2 and randoms[0] == randoms[1]:
        fails += 1
        print("  FAIL genSecRandomBytes: two calls gave the same bytes")
    total = len([c for c in checks if c[0] != "speed"])
    print("crypto: %d of %d checks as libsodium (PyNaCl), hashlib, cryptography and crc.cpp give them; %.1f s"
          % (total - fails, total, dt))
    print("crypto speed (ms per operation, test build: bounds checks, no optimisation): " + speed.replace("\t", ", "))
    print("crypto exactness: the largest magnitude sodium_25519_c's Doubles reached over these checks: %.0f (2^%.1f; "
          "exact below 2^53)" % (limb, math.log2(limb) if limb > 0 else 0))
    if not keep:
        shutil.rmtree(WORK, ignore_errors=True)
    return fails == 0


def main(argv):
    stage = argv[argv.index("--stage") + 1] if "--stage" in argv else None
    out = argv[argv.index("--out") + 1] if "--out" in argv else None
    if out:
        os.makedirs(out, exist_ok=True)
    return 0 if run("--keep" in argv, stage, out) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
