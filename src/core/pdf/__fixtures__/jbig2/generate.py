# Writes jbig2.json: JBIG2 streams with the pages a reference decoder makes of them.
#
# Run inside Debian trixie, which packages the jbig2enc encoder (`jbig2`) and the jbig2dec
# reference decoder (see README.md):
#   docker run --rm -v "$PWD:/work" -w /work debian:trixie bash -c \
#     "apt-get update -qq && apt-get install -y -qq jbig2 jbig2dec python3-pil && python3 generate.py"
import base64
import hashlib
import json
import os
import subprocess
import tempfile

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = tempfile.mkdtemp()


def text_image():
    image = Image.new('1', (203, 46), 1)
    draw = ImageDraw.Draw(image)
    draw.text((5, 3), 'JBIG2 generic 0123', fill=0)
    draw.text((5, 20), 'region, AT pixels', fill=0)
    draw.rectangle((170, 4, 196, 40), outline=0)
    draw.line((0, 45, 202, 30), fill=0)
    return image


def page_image():
    image = Image.new('1', (240, 90), 1)
    draw = ImageDraw.Draw(image)
    for line in range(6):
        draw.text((4, 4 + line * 14), 'the quick brown fox %d jumps' % line, fill=0)
    return image


def save(image, name):
    path = os.path.join(WORK, name + '.png')
    image.save(path)
    return path


def run(*args):
    return subprocess.run(args, check=True, capture_output=True, cwd=WORK).stdout


def read(path):
    with open(path, 'rb') as file:
        return file.read()


def pbm_pages(data):
    """Splits concatenated binary PBM (P4) files into (width, height, packed rows)."""
    pages = []
    at = 0
    while at < len(data):
        fields = []
        while len(fields) < 3:
            while data[at:at + 1].isspace():
                at += 1
            if data[at:at + 1] == b'#':
                at = data.index(b'\n', at) + 1
                continue
            end = at
            while not data[end:end + 1].isspace():
                end += 1
            fields.append(data[at:end])
            at = end
        at += 1
        width, height = int(fields[1]), int(fields[2])
        size = (width + 7) // 8 * height
        pages.append((width, height, data[at:at + size]))
        at += size
    return pages


def reference(paths, embedded):
    out = os.path.join(WORK, 'out.pbm')
    if os.path.exists(out):
        os.remove(out)
    run('jbig2dec', '-q', *(['-e'] if embedded else []), '-t', 'pbm', '-o', out, *paths)
    pages = []
    for width, height, bits in pbm_pages(read(out)):
        page = {'width': width, 'height': height}
        if len(bits) <= 4096:
            page['bits'] = base64.b64encode(bits).decode()
        else:
            page['sha256'] = hashlib.sha256(bits).hexdigest()
        pages.append(page)
    return pages


def case(name, data, pages, globals_data=None, embedded=True, source='jbig2enc 0.30'):
    item = {'name': name, 'source': source, 'embedded': embedded}
    if globals_data is not None:
        item['globals'] = base64.b64encode(globals_data).decode()
    item['data'] = base64.b64encode(data).decode()
    item['pages'] = pages
    return item


def embedded_case(name, data_path, globals_path=None, source='jbig2enc 0.30'):
    paths = [globals_path, data_path] if globals_path else [data_path]
    globals_data = read(globals_path) if globals_path else None
    return case(name, read(data_path), reference(paths, True), globals_data, True, source)


cases = []
text = save(text_image(), 'text')
page = save(page_image(), 'page')

for label, flags in (('generic', []), ('generic tpgdon', ['-d'])):
    path = os.path.join(WORK, label.replace(' ', '-') + '.jb2')
    with open(path, 'wb') as file:
        file.write(run('jbig2', '-p', *flags, text))
    cases.append(embedded_case(label, path))

# jbig2enc's refinement (-r) is disabled ("refinement broke"); Annex H covers refinement.
run('jbig2', '-s', '-p', '-b', 'symbol', page)
cases.append(embedded_case('symbol', os.path.join(WORK, 'symbol.0000'),
                           os.path.join(WORK, 'symbol.sym')))

# A whole file (with the file header) in sequential organisation.
path = os.path.join(WORK, 'file.jb2')
with open(path, 'wb') as file:
    file.write(run('jbig2', page))
cases.append(case('file', read(path), reference([path], False), embedded=False))

annex = os.path.join(HERE, 'annex-h.jbig2')
cases.append(case('annex h', read(annex), reference([annex], False), embedded=False,
                  source='ITU-T T.88 Annex H.2 test bitstream'))
cases.append(embedded_case('pdfbox issue 21', os.path.join(HERE, '21.jb2'),
                           os.path.join(HERE, '21.glob'), source='apache/pdfbox-jbig2'))

with open(os.path.join(HERE, 'jbig2.json'), 'w') as file:
    json.dump(cases, file, indent=1)
