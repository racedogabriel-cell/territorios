import os
import struct
import zlib

def chunk(tag, data):
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

def inside_round(x, y, left, top, right, bot, rad):
    if x < left or x >= right or y < top or y >= bot:
        return False
    if x < left + rad and y < top + rad:
        dx, dy = x - (left + rad), y - (top + rad)
        return dx * dx + dy * dy <= rad * rad
    if x >= right - rad and y < top + rad:
        dx, dy = x - (right - rad - 1), y - (top + rad)
        return dx * dx + dy * dy <= rad * rad
    if x < left + rad and y >= bot - rad:
        dx, dy = x - (left + rad), y - (bot - rad - 1)
        return dx * dx + dy * dy <= rad * rad
    if x >= right - rad and y >= bot - rad:
        dx, dy = x - (right - rad - 1), y - (bot - rad - 1)
        return dx * dx + dy * dy <= rad * rad
    return True

def png(size):
    raw = bytearray()
    margin = size * 0.20
    left, top = margin, size * 0.30
    right, bot = size - margin, size * 0.70
    rad = size * 0.08
    gold_t, gold_b = size * 0.62, size * 0.68
    gold_l, gold_r = size * 0.34, size * 0.66
    for y in range(size):
        raw.append(0)
        for x in range(size):
            if inside_round(x + 0.5, y + 0.5, left, top, right, bot, rad):
                if gold_l <= x < gold_r and gold_t <= y < gold_b:
                    raw.extend((183, 139, 30, 255))
                else:
                    raw.extend((255, 255, 255, 255))
            else:
                raw.extend((44, 62, 80, 255))
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    sig = bytes([137, 80, 78, 71, 13, 10, 26, 10])
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b"")

def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for folder in ("docs", "Netlify"):
        dest = os.path.join(root, folder)
        for size, name in ((192, "icon-192.png"), (512, "icon-512.png")):
            path = os.path.join(dest, name)
            with open(path, "wb") as fh:
                fh.write(png(size))
            print(path, os.path.getsize(path))

if __name__ == "__main__":
    main()
