"""
Genera iconos PNG reales para el Office Add-in.
"""
import struct
import zlib
import os

STATIC_DIR = os.path.join(os.path.dirname(__file__), 'static')


def create_png(width, height, color=(33, 115, 70)):
    """Crea un PNG simple con un color sólido y texto 'IA'"""
    r, g, b = color

    # Crear pixels (RGBA)
    pixels = []
    for y in range(height):
        row = b'\x00'  # filter byte
        for x in range(width):
            # Fondo verde con bordes redondeados (simplificado como cuadrado)
            row += struct.pack('BBBB', r, g, b, 255)
        pixels.append(row)

    raw_data = b''.join(pixels)

    def make_chunk(chunk_type, data):
        chunk = chunk_type + data
        return struct.pack('>I', len(data)) + chunk + struct.pack('>I', zlib.crc32(chunk) & 0xffffffff)

    # PNG signature
    signature = b'\x89PNG\r\n\x1a\n'

    # IHDR
    ihdr_data = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)
    ihdr = make_chunk(b'IHDR', ihdr_data)

    # IDAT
    compressed = zlib.compress(raw_data)
    idat = make_chunk(b'IDAT', compressed)

    # IEND
    iend = make_chunk(b'IEND', b'')

    return signature + ihdr + idat + iend


sizes = {
    'icon-16.png': 16,
    'icon-32.png': 32,
    'icon-80.png': 80,
}

for filename, size in sizes.items():
    path = os.path.join(STATIC_DIR, filename)
    png_data = create_png(size, size)
    with open(path, 'wb') as f:
        f.write(png_data)
    print("  Creado: %s (%dx%d)" % (filename, size, size))

print("  Iconos PNG generados correctamente.")
