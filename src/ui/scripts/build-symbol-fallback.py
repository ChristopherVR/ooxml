"""Generate the offline symbol font from the pinned, separately downloaded source.

Requires fonttools[woff] 4.65.0. See PROVENANCE.md for source and licence.
No downloads occur in this script or in the shipped renderer.
"""
import base64
import hashlib
from io import BytesIO
from pathlib import Path
import sys

from fontTools import subset
from fontTools.ttLib import TTFont

source = Path(sys.argv[1])
digest = hashlib.sha256(source.read_bytes()).hexdigest()
assert digest == '7d5fb73b7ca67a6798101741f5d280a3d016a56a197afcd4199dbb57b4b82a21'
font = TTFont(source)
options = subset.Options()
options.flavor = 'woff2'
subsetter = subset.Subsetter(options=options)
subsetter.populate(unicodes=range(0x2000, 0x2c00))
subsetter.subset(font)
assert all(ord(c) in font.getBestCmap() for c in '■▪❑❖⮚✓✔➢•')
font.flavor = 'woff2'
buffer = BytesIO()
font.save(buffer)
encoded = base64.b64encode(buffer.getvalue()).decode('ascii')
destination = Path(__file__).resolve().parents[1] / 'src/fonts/assets/symbol-fallback-data.ts'
destination.write_text(
    '// Noto Sans Symbols 2, U+2000-2BFF subset. SIL OFL 1.1, see OFL-NotoSansSymbols2.txt.\n'
    f'// Source SHA-256: {digest}\n'
    'export const SYMBOL_FALLBACK_DATA =\n'
    f"\t'data:font/woff2;base64,{encoded}';\n",
    encoding='utf-8',
)
