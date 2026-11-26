import { first } from '../xml/index';
import { parseCssDeclarations } from '../text/html-tokens';
import type { InlineImage, PicturePlacement } from './inline-model';
import { isVmlPictureWashout } from './picture-watermark';
import type { XmlElement } from './xml';

const WORD10_NS = 'urn:schemas-microsoft-com:office:word';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const PX_PER_UNIT = new Map([
	['px', 1],
	['pt', 96 / 72],
	['in', 96],
	['cm', 96 / 2.54],
	['mm', 96 / 25.4],
]);

// MS-OI29500, Part 4, 19.1.2.19.
const HORIZONTAL = new Map<string, PicturePlacement['relativeFrom']>([
	['margin', 'margin'],
	['page', 'page'],
	['text', 'column'],
	['char', 'character'],
	['left-margin', 'leftMargin'],
	['right-margin', 'rightMargin'],
	['inner-margin-area', 'insideMargin'],
	['outer-margin-area', 'outsideMargin'],
]);
const VERTICAL = new Map<string, PicturePlacement['relativeFromV']>([
	['margin', 'margin'],
	['page', 'page'],
	['text', 'paragraph'],
	['line', 'line'],
	['top-margin', 'topMargin'],
	['bottom-margin', 'bottomMargin'],
	['inner-margin-area', 'insideMargin'],
	['outer-margin-area', 'outsideMargin'],
]);

function lengthPx(value: string | undefined): number {
	const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(px|pt|in|cm|mm)?$/.exec(value ?? '');
	const px = Number(match?.[1]) * (PX_PER_UNIT.get(match?.[2] ?? 'px') ?? 1);
	return Number.isFinite(px) ? px : 0;
}

/** The model fields of a resolved absolute VML picture. */
export function vmlFloat(
	shape: XmlElement,
	data: XmlElement,
):
	| Pick<InlineImage, 'widthPx' | 'heightPx' | 'anchored' | 'placement' | 'pictureFill'>
	| undefined {
	const style = parseCssDeclarations(shape.getAttribute('style') ?? '');
	const property = (name: string) => style.get(name)?.toLowerCase();
	if (property('position') !== 'absolute') return undefined;
	const wrapNode = first(shape, 'wrap', WORD10_NS);
	const type = wrapNode?.getAttribute('type');
	const side = wrapNode?.getAttribute('side');
	const h = property('mso-position-horizontal');
	const v = property('mso-position-vertical');
	const relativeFrom = HORIZONTAL.get(property('mso-position-horizontal-relative') ?? '');
	const relativeFromV = VERTICAL.get(property('mso-position-vertical-relative') ?? '');
	const placement: PicturePlacement = {
		relativeFrom:
			relativeFrom ?? HORIZONTAL.get(wrapNode?.getAttribute('anchorx') ?? '') ?? 'column',
		relativeFromV:
			relativeFromV ?? VERTICAL.get(wrapNode?.getAttribute('anchory') ?? '') ?? 'paragraph',
		...(h === 'left' || h === 'center' || h === 'right' || h === 'inside' || h === 'outside'
			? { align: h }
			: { offsetXPx: lengthPx(property('margin-left')) }),
		...(v === 'top' || v === 'center' || v === 'bottom' || v === 'inside' || v === 'outside'
			? { alignV: v }
			: { offsetYPx: lengthPx(property('margin-top')) }),
		...(Number(property('z-index')) < 0 ? { behindText: true } : {}),
		wrap:
			type === 'square' || type === 'tight' || type === 'through' || type === 'topAndBottom'
				? type
				: 'none',
		...(side === 'both'
			? { wrapText: 'bothSides' }
			: side === 'left' || side === 'right' || side === 'largest'
				? { wrapText: side }
				: {}),
		...(/^(?:f|false|0)$/.test(shape.getAttributeNS(OFFICE_NS, 'allowincell') ?? '')
			? { layoutInCell: false }
			: {}),
	};
	return {
		widthPx: Math.max(0, lengthPx(property('width'))),
		heightPx: Math.max(0, lengthPx(property('height'))),
		anchored: true,
		placement,
		...(isVmlPictureWashout(data)
			? {
					// Match the watermark painter's linear washout, channel x -> 0.09x + 0.805.
					pictureFill: {
						kind: 'picture',
						effects: [{ kind: 'luminance', brightness: 70 / 109, contrast: -0.91 }],
					},
				}
			: {}),
	};
}
