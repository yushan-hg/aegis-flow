"""Vẽ icon app (không cần thư viện ảnh): nền tím đậm + hình bản (pallet) xếp hàng. Chạy: python tools/make_icons.py"""
import zlib, struct, os

def png(path, size, draw):
    rows = []
    for y in range(size):
        row = bytearray([0])
        for x in range(size):
            row += bytes(draw(x / size, y / size))
        rows.append(bytes(row))
    raw = b''.join(rows)
    def chunk(t, d):
        c = struct.pack('>I', len(d)) + t + d
        return c + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    data = b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', size, size, 8, 6, 0, 0, 0)) \
         + chunk(b'IDAT', zlib.compress(raw, 9)) + chunk(b'IEND', b'')
    open(path, 'wb').write(data)

def rect(u, v, x0, y0, x1, y1):
    return x0 <= u <= x1 and y0 <= v <= y1

def draw(u, v):
    # nền: chuyển màu chéo tím đậm → xanh tím
    t = (u + v) / 2
    r, g, b = int(26 + 40 * t), int(30 + 38 * t), int(70 + 80 * t)
    # 3 hàng thùng xếp trên một tấm pallet
    box = (236, 238, 252)
    for row in range(3):
        for col in range(3):
            x0 = 0.22 + col * 0.19; y0 = 0.20 + row * 0.15
            if rect(u, v, x0, y0, x0 + 0.165, y0 + 0.125):
                return (box[0], box[1], box[2], 255)
    # tấm pallet (3 thanh) + hai chân
    amber = (255, 176, 32)
    if rect(u, v, 0.18, 0.68, 0.82, 0.74): return (*amber, 255)
    for x0 in (0.18, 0.46, 0.76):
        if rect(u, v, x0, 0.74, x0 + 0.06, 0.82): return (*amber, 255)
    return (r, g, b, 255)

here = os.path.dirname(os.path.abspath(__file__))
out = os.path.join(here, '..')
png(os.path.join(out, 'icon-512.png'), 512, draw)
png(os.path.join(out, 'icon-192.png'), 192, draw)
png(os.path.join(out, 'apple-touch-icon.png'), 180, draw)
print('ok')
