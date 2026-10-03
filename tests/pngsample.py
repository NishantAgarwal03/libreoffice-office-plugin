"""Minimal pure-Python PNG reader for tests (LibreOffice's bundled Python has no Pillow).

Used to *see* what LibreOffice renders: export a sheet to PNG and sample cell colours.
Supports non-interlaced 8-bit greyscale / RGB / RGBA / palette PNGs (what calc_png_Export writes).
"""

import struct
import zlib


class Png:
    def __init__(self, width, height, channels, rows, palette=None):
        self.width, self.height, self.channels, self.rows, self.palette = width, height, channels, rows, palette

    def pixel(self, x, y):
        x = max(0, min(self.width - 1, x))
        y = max(0, min(self.height - 1, y))
        row, c = self.rows[y], self.channels
        if self.palette is not None:
            return self.palette[row[x]]
        px = row[x * c:(x + 1) * c]
        if c == 1:
            return (px[0],) * 3
        return tuple(px[:3])


def read_png(path):
    with open(path, "rb") as fh:
        data = fh.read()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")
    pos, idat, palette, header = 8, b"", None, None
    while pos < len(data):
        length, kind = struct.unpack(">I4s", data[pos:pos + 8])
        body = data[pos + 8:pos + 8 + length]
        pos += 12 + length
        if kind == b"IHDR":
            header = struct.unpack(">IIBBBBB", body)
        elif kind == b"PLTE":
            palette = [tuple(body[i:i + 3]) for i in range(0, len(body), 3)]
        elif kind == b"IDAT":
            idat += body
        elif kind == b"IEND":
            break
    width, height, depth, color_type, _, _, interlace = header
    if depth != 8 or interlace != 0:
        raise ValueError("unsupported PNG (depth %d, interlace %d)" % (depth, interlace))
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color_type]
    raw = zlib.decompress(idat)
    stride, rows, prev = width * channels, [], bytearray(width * channels)
    i = 0
    for _ in range(height):
        ftype = raw[i]
        line = bytearray(raw[i + 1:i + 1 + stride])
        i += 1 + stride
        for x in range(stride):
            a = line[x - channels] if x >= channels else 0
            b = prev[x]
            c = prev[x - channels] if x >= channels else 0
            if ftype == 1:
                line[x] = (line[x] + a) & 255
            elif ftype == 2:
                line[x] = (line[x] + b) & 255
            elif ftype == 3:
                line[x] = (line[x] + ((a + b) >> 1)) & 255
            elif ftype == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                pred = a if pa <= pb and pa <= pc else (b if pb <= pc else c)
                line[x] = (line[x] + pred) & 255
        rows.append(line)
        prev = line
    return Png(width, height, channels, rows, palette if color_type == 3 else None)
