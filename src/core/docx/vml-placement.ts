import { first, makeNS } from '../xml/index';
import { parseCssDeclarations } from '../text/html-tokens';
import type { InlineImage, PicturePlacement } from './inline-model';
import { isVmlPictureWashout } from './picture-watermark';
import type { XmlElement } from './xml';

const WORD10_NS = 'urn:schemas-microsoft-com:office:word';
export const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const PX_PER_UNIT = new Map([
	['px', 1],
	['pt', 96 / 72],
	['in', 96],
	['cm', 96 / 2.54],
	['mm', 96 / 25.4],
]);

// MS-OI29500, Part 4, 19.1.2.19.
const HORIZONTAL = new Map<string, NonNullable<PicturePlacement['relativeFrom']>>([
	['margin', 'margin'],
	['page', 'page'],
	['text', 'column'],
	['char', 'character'],
	['left-margin', 'leftMargin'],
	['right-margin', 'rightMargin'],
	['inner-margin-area', 'insideMargin'],
	['outer-margin-area', 'outsideMargin'],
]);
const VERTICAL = new Map<string, NonNullable<PicturePlacement['relativeFromV']>>([
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

// VML uses the sign for the text layer. Shift the negative range without reversing its order.
const Z_INDEX_RANGE = 2 ** 31;

/** The source shape used for both loading and patching a VML picture. */
export function vmlPictureShape(pict: XmlElement): XmlElement | undefined {
	return first(pict, 'shape', VML_NS) ?? pict.getElementsByTagNameNS(VML_NS, 'shape')[0];
}

/** The model fields of a resolved absolute VML picture. */
export function vmlFloat(
	shape: XmlElement,
	data: XmlElement,
):
	| Pick<
			InlineImage,
			'widthPx' | 'heightPx' | 'anchored' | 'placement' | 'pictureFill' | 'vmlFloat'
	  >
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
	const zIndex = Number(property('z-index'));
	const hasZIndex = Number.isInteger(zIndex) && zIndex >= -Z_INDEX_RANGE && zIndex < Z_INDEX_RANGE;
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
		...(zIndex < 0 ? { behindText: true } : {}),
		...(hasZIndex ? { relativeHeight: zIndex < 0 ? zIndex + Z_INDEX_RANGE : zIndex } : {}),
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
		vmlFloat: true,
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

/** Patch just the changed declarations, keeping the other tokens and their order. */
function patchStyle(style: string, changes: Map<string, string>): string {
	const declarations = style.split(';').map((declaration) => {
		const colon = declaration.indexOf(':');
		if (colon < 0) return declaration;
		const name = declaration.slice(0, colon).trim().toLowerCase();
		const value = changes.get(name);
		return value === undefined ? declaration : `${declaration.slice(0, colon + 1)}${value}`;
	});
	const source = parseCssDeclarations(style);
	for (const [name, value] of changes) {
		if (source.has(name)) continue;
		const end = declarations.length - 1;
		declarations.splice(declarations[end]?.trim() === '' ? end : end + 1, 0, `${name}:${value}`);
	}
	return declarations.join(';');
}

function vmlReference<T extends string>(map: ReadonlyMap<string, T>, value: T | undefined): string {
	for (const [vml, placement] of map) if (placement === value) return vml;
	throw new RangeError(`Unsupported VML picture reference frame: ${value}`);
}

const points = (px: number): string => `${(px * 0.75).toFixed(6).replace(/\.?0+$/, '')}pt`;

/** Writes the geometry, placement, wrapping and stacking edits of an existing VML float. */
export function patchVmlFloat(pict: XmlElement, image: InlineImage, base: InlineImage): void {
	const shape = vmlPictureShape(pict);
	const previous = base.placement;
	const next = image.placement;
	if (!shape || !previous) return;
	if (!image.anchored || !next)
		throw new Error('Inline conversion is unavailable for legacy VML pictures.');
	const changes = new Map<string, string>();
	if (image.widthPx !== base.widthPx) changes.set('width', points(image.widthPx));
	if (image.heightPx !== base.heightPx) changes.set('height', points(image.heightPx));
	if (next.align !== previous.align)
		changes.set('mso-position-horizontal', next.align ?? 'absolute');
	if (!next.align && (next.offsetXPx !== previous.offsetXPx || previous.align))
		changes.set('margin-left', points(next.offsetXPx ?? 0));
	if (next.alignV !== previous.alignV)
		changes.set('mso-position-vertical', next.alignV ?? 'absolute');
	if (!next.alignV && (next.offsetYPx !== previous.offsetYPx || previous.alignV))
		changes.set('margin-top', points(next.offsetYPx ?? 0));
	const horizontal =
		next.relativeFrom !== previous.relativeFrom
			? vmlReference(HORIZONTAL, next.relativeFrom ?? 'column')
			: undefined;
	const vertical =
		next.relativeFromV !== previous.relativeFromV
			? vmlReference(VERTICAL, next.relativeFromV ?? 'paragraph')
			: undefined;
	if (horizontal) changes.set('mso-position-horizontal-relative', horizontal);
	if (vertical) changes.set('mso-position-vertical-relative', vertical);
	if (
		Boolean(next.behindText) !== Boolean(previous.behindText) ||
		next.relativeHeight !== previous.relativeHeight
	) {
		const order = Math.max(0, Math.min(Z_INDEX_RANGE - 1, Math.round(next.relativeHeight ?? 0)));
		changes.set('z-index', String(next.behindText ? order - Z_INDEX_RANGE : order));
	}
	if (changes.size)
		shape.setAttribute('style', patchStyle(shape.getAttribute('style') ?? '', changes));

	const wrapChanged = next.wrap !== previous.wrap || next.wrapText !== previous.wrapText;
	const existingWrap = first(shape, 'wrap', WORD10_NS);
	if (!wrapChanged && !(existingWrap && (horizontal || vertical))) return;
	const wrap = existingWrap ?? makeNS(pict.ownerDocument, WORD10_NS, 'w10:wrap');
	if (next.wrap !== previous.wrap) wrap.setAttribute('type', next.wrap);
	if (next.wrapText !== previous.wrapText) {
		if (next.wrapText)
			wrap.setAttribute('side', next.wrapText === 'bothSides' ? 'both' : next.wrapText);
		else wrap.removeAttribute('side');
	}
	if (horizontal)
		wrap.setAttribute(
			'anchorx',
			/^(margin|page|text|char)$/.test(horizontal) ? horizontal : 'margin',
		);
	if (vertical)
		wrap.setAttribute('anchory', /^(margin|page|text|line)$/.test(vertical) ? vertical : 'margin');
	if (!existingWrap) shape.appendChild(wrap);
}
