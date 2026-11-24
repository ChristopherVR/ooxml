# Writes ccitt.json: small 1-bit images encoded by libtiff (through Pillow) as CCITT Group 4,
# Group 3 one-dimensional (with end-of-line codes), Group 3 two-dimensional and CCITT RLE
# (byte-aligned Modified Huffman), with the expected pixels. Run: python ccitt.py
import base64
import io
import json
import os

from PIL import Image, ImageDraw, TiffImagePlugin


def sample(width, height):
    image = Image.new('1', (width, height), 1)
    draw = ImageDraw.Draw(image)
    draw.text((4, 2), 'CCITT fax 0123', fill=0)
    draw.rectangle((width - 30, 3, width - 6, height - 4), outline=0)
    draw.line((0, height - 1, width - 1, 0), fill=0)
    for x in range(0, width, 7):
        draw.point((x, height - 2), fill=0)
    return image


def wide():
    # Runs longer than 1728 use the extended make-up codes.
    image = Image.new('1', (2600, 6), 1)
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 1, 1999, 2), fill=0)
    draw.rectangle((2010, 3, 2599, 4), fill=0)
    return image


def strip(image, compression, options=None):
    # libtiff codes 0 bits as white runs; Pillow stores white as 1, so encode the inverse.
    image = image.convert('L').point(lambda value: 255 - value).convert('1')
    out = io.BytesIO()
    info = TiffImagePlugin.ImageFileDirectory_v2()
    if options is not None:
        info[292] = options
    image.save(out, 'TIFF', compression=compression, tiffinfo=info)
    out.seek(0)
    tiff = Image.open(out)
    offsets = tiff.tag_v2[273]
    counts = tiff.tag_v2[279]
    data = out.getvalue()
    return b''.join(data[o:o + c] for o, c in zip(offsets, counts))


def packed(image):
    # PDF CCITT output: one bit per pixel, 1 for white (BlackIs1 false), rows padded with white.
    data = bytearray(image.tobytes())
    stride = (image.width + 7) // 8
    pad = stride * 8 - image.width
    if pad:
        for row in range(image.height):
            data[row * stride + stride - 1] |= (1 << pad) - 1
    return bytes(data)


cases = []
for name, image in (('sample', sample(150, 40)), ('wide', wide())):
    expected = base64.b64encode(packed(image)).decode()
    for label, compression, options, k, eol, align in (
        ('group4', 'group4', None, -1, False, False),
        ('group3-1d', 'group3', 0, 0, True, False),
        ('group3-2d', 'group3', 1, 4, True, False),
        ('rle', 'tiff_ccitt', None, 0, False, True),
    ):
        cases.append({
            'name': f'{name} {label}',
            'width': image.width,
            'height': image.height,
            'k': k,
            'endOfLine': eol,
            'encodedByteAlign': align,
            'data': base64.b64encode(strip(image, compression, options)).decode(),
            'expected': expected,
        })

with open(os.path.join(os.path.dirname(__file__), 'ccitt.json'), 'w') as file:
    json.dump(cases, file, indent=1)
