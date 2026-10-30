"""The loudness of a captured WAV over time: the RMS and peak of each window, and the stretches with sound.
python -I tools\\wav_activity.py <file.wav> [window seconds]"""
import sys
import wave

import numpy as np

path = sys.argv[1]
win = float(sys.argv[2]) if len(sys.argv) > 2 else 0.25
w = wave.open(path, "rb")
rate, ch, n = w.getframerate(), w.getnchannels(), w.getnframes()
data = np.frombuffer(w.readframes(n), dtype="<i2").reshape(-1, ch).astype(np.float64)
print("%s: %d Hz, %d channels, %.1f s, peak %d" % (path, rate, ch, n / rate, int(np.abs(data).max()) if n else 0))
step = int(rate * win)
active = []
for i in range(0, n, step):
    blk = data[i:i + step]
    rms = np.sqrt((blk ** 2).mean(axis=0))
    if rms.max() > 30:
        active.append((i / rate, rms))
prev = None
for t, rms in active:
    if prev is None or t - prev > win * 1.5:
        print("  sound from %6.2f s:" % t, end="")
    print(" %d/%d" % (rms[0], rms[-1]), end="")
    prev = t
print()
print("%d of %d windows with sound" % (len(active), (n + step - 1) // step))
