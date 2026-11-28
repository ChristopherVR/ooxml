/**
 * The header's watermark for the Word PDF export, behind the page's text as Print Layout draws it
 * (`watermarkElement` in `print-header-footer.ts`): text as real, bold text centred on the page,
 * turned 45 degrees when diagonal and at half opacity (a constant-alpha graphics state) when
 * semitransparent; a picture centred on the margins, washed out as Word's Washout does. DOM-free.
 */
import { fontMetrics, type LayoutPageBox, type TextMeasurer } from '../layout/index';
import { watermarkSizePt, type PictureWatermarkSpec, type WatermarkSpec } from '../index';
import { WASHOUT_INTERCEPT, WASHOUT_SLOPE } from '../picture-watermark';
import type { PdfImageResource } from '../../pdf/write/pdf-vector';
import type { DocxPdfPainter } from './pdf-blocks';
import type { DocxPdfStoryGraphics } from './pdf-story-graphics';

const DEFAULT_FAMILY = 'Calibri';
const DEFAULT_COLOR = '#c0c0c0';
const washed = new WeakMap<PdfImageResource, PdfImageResource>();

function washedOut(resource: PdfImageResource): PdfImageResource | undefined {
	if (resource.filter) return undefined;
	let result = washed.get(resource);
	if (!result) {
		const data = new Uint8Array(resource.data.length);
		for (let index = 0; index < data.length; index++)
			data[index] = Math.min(
				255,
				Math.round(resource.data[index]! * WASHOUT_SLOPE + WASHOUT_INTERCEPT * 255),
			);
		result = { ...resource, data };
		washed.set(resource, result);
	}
	return result;
}

function paintText(
	painter: DocxPdfPainter,
	spec: WatermarkSpec,
	page: LayoutPageBox,
	measurer: TextMeasurer,
): void {
	const shapeSize =
		spec.widthPt && spec.heightPt
			? { width: spec.widthPt, height: spec.heightPt }
			: watermarkSizePt(spec);
	const targetWidthPx = (shapeSize.width * 4) / 3;
	const targetHeightPx = (shapeSize.height * 4) / 3;
	const sizePx = Math.max(12, Math.min(220, targetHeightPx));
	const family = spec.fontFamily ?? DEFAULT_FAMILY;
	const font = { family, sizePx, bold: true };
	const widthPx = Math.max(1, measurer.widthOf(spec.text, font));
	const { ascent, descent } = fontMetrics(measurer, font);
	const heightPx = Math.max(1, ascent + descent);
	const scaleX = targetWidthPx / widthPx;
	const scaleY = targetHeightPx / heightPx;
	const pdf = painter.page;
	pdf.save();
	if (spec.semitransparent) pdf.alpha(0.5);
	const turn = spec.layout === 'diagonal' ? Math.SQRT1_2 : 0;
	const keep = spec.layout === 'diagonal' ? Math.SQRT1_2 : 1;
	// About the page's centre, a quarter of a half turn anticlockwise (CSS `rotate(-45deg)`),
	// then stretched to fill the watermark shape (`v:textpath fitshape="t"`).
	pdf.concat([keep, -turn, turn, keep, page.widthPx / 2, page.heightPx / 2]);
	pdf.concat([scaleX, 0, 0, scaleY, 0, 0]);
	const fragment = {
		text: spec.text,
		xPx: 0,
		widthPx,
		runIndex: 0,
		fontFamily: family,
		fontSizePt: (sizePx * 72) / 96,
		bold: true,
		color: /^#[0-9a-f]{6}$/i.test(spec.color) ? spec.color : DEFAULT_COLOR,
	};
	const line = {
		yPx: 0,
		heightPx,
		baselinePx: ascent,
		fragments: [fragment],
		sourceStart: 0,
		sourceEnd: 0,
	};
	painter.text.draw(pdf, fragment, -widthPx / 2, line, -heightPx / 2);
	pdf.restore();
}

function paintPicture(painter: DocxPdfPainter, spec: PictureWatermarkSpec, page: LayoutPageBox) {
	const widthPx = (spec.widthPt * 4) / 3;
	const heightPx = (spec.heightPt * 4) / 3;
	const areaWidth = page.widthPx - page.marginLeftPx - page.marginRightPx;
	const areaHeight = page.heightPx - page.marginTopPx - page.marginBottomPx;
	const x = page.marginLeftPx + (areaWidth - widthPx) / 2;
	const y = page.marginTopPx + (areaHeight - heightPx) / 2;
	const picture = { partName: spec.partName, contentType: spec.contentType, widthPx, heightPx };
	const { pictures, page: pdf } = painter;
	const resource = pictures.resource(spec.partName);
	const light = spec.washout && resource ? washedOut(resource) : undefined;
	if (light) {
		pdf.save();
		pdf.concat([widthPx, 0, 0, -heightPx, x, y + heightPx]);
		pdf.drawImage(light);
		pdf.restore();
		return;
	}
	pictures.draw(pdf, picture, x, y);
	if (!spec.washout || !resource) return;
	// A JPEG is written as it is: a grey layer in Screen mode gives (1 - g) x + g, with g the
	// intercept, so black rises to it and white stays white (slope 0.2 for Word's 0.3).
	pdf.save();
	pdf.alpha(1, 1, 'Screen');
	pdf.fillColor(WASHOUT_INTERCEPT, WASHOUT_INTERCEPT, WASHOUT_INTERCEPT);
	pdf.rect(x, y, widthPx, heightPx);
	pdf.op('f');
	pdf.restore();
}

/** Draws the page's watermark, if its header has one. */
export function paintWatermark(
	painter: DocxPdfPainter,
	graphics: DocxPdfStoryGraphics | undefined,
	page: LayoutPageBox,
	measurer: TextMeasurer,
): void {
	if (graphics?.pictureWatermark) paintPicture(painter, graphics.pictureWatermark, page);
	else if (graphics?.watermark) paintText(painter, graphics.watermark, page, measurer);
}
