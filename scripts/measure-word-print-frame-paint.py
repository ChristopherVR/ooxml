"""Verify Chromium nested-frame outside stroke and visible ink from optional Playwright evidence."""
import json,sys
from pathlib import Path
import fitz
root=Path(sys.argv[1])
data=json.loads((root/'word-mixed-frame-geometry.json').read_text(encoding='utf-8'))
image=fitz.Pixmap(str(root/'word-mixed-frame.png'))
scale=image.width/data['viewport']['width']
frame,body,ink=(data[key] for key in ('frame','body','ink'))
x=int((frame['x']+frame['width']-1)*scale);y=int((frame['y']+8)*scale)
color=image.pixel(x,y)[:3]
dark=sum(max(image.pixel(xx,yy)[:3])<80 for yy in range(int(ink['y']*scale),int((ink['y']+ink['height'])*scale)) for xx in range(int(ink['x']*scale),int((ink['x']+min(ink['width'],90))*scale)))
assert x/scale>body['x']+body['width'], 'Stroke probe must be outside the source cell'
assert color==(255,0,0), f'Outside red border was clipped: {color}'
assert dark>20, f'Projected text was hidden by frame shading: {dark} dark pixels'
clipped=fitz.Pixmap(str(root/'word-mixed-frame-clipped.png'))
clip_probe_y=int((frame['y']+14)*scale)
clip_color=clipped.pixel(x,clip_probe_y)[:3]
assert clip_color!=(255,0,0), f'Frame stroke escaped an exact-height vertical clip: {clip_color}'

print(json.dumps({'outsideStrokePixel':[x,y],'outsideStrokeRGB':color,'visibleDarkInkPixels':dark,'screenshotScale':scale,'belowClipPixelRGB':clip_color}))
