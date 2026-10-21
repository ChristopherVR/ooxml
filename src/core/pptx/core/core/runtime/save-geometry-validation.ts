import type { PptxElement, PptxSlide } from '../../types';
import { findNonVisualProps } from '../../utils/placeholder-xml';
import { xmlChild, xmlHasChild } from '../../utils/xml-access';

const FRAME_FIELDS = ['x', 'y', 'width', 'height'] as const;
const OPTIONAL_FIELDS = [
	'xEmu',
	'yEmu',
	'widthEmu',
	'heightEmu',
	'rotation',
	'skewX',
	'skewY',
	'chOffXEmu',
	'chOffYEmu',
	'chExtWidthEmu',
	'chExtHeightEmu',
] as const;

/** An unchanged placeholder can omit its own frame, including absent resolved axes. */
function omitsInheritedFrame(element: PptxElement): boolean {
	if (!['text', 'shape', 'picture', 'media'].includes(element.type) || !element.rawXml)
		return false;
	if (!xmlHasChild(findNonVisualProps(element.rawXml), 'p:ph')) return false;
	if (
		xmlHasChild(xmlChild(element.rawXml, 'p:spPr'), 'a:xfrm') ||
		xmlHasChild(element.rawXml, 'p:xfrm')
	)
		return false;
	const baseline = element.inheritedTransform;
	if (!baseline) return FRAME_FIELDS.every((key) => element[key] === undefined);
	return (
		FRAME_FIELDS.every(
			(key) =>
				Number.isFinite(baseline[key]) &&
				(element[key] === undefined || element[key] === baseline[key]),
		) &&
		(['rotation', 'skewX', 'skewY', 'flipHorizontal', 'flipVertical'] as const).every(
			(key) => (element[key] ?? 0) === (baseline[key] ?? 0),
		)
	);
}

/** Reject invalid geometry before save mutates relationships, XML, or archive parts. */
export function validateElementGeometry(
	element: PptxElement,
	context: string,
	emuPerPx = 9525,
): void {
	const inherited = omitsInheritedFrame(element);
	const record = element as unknown as Record<string, unknown>;
	const fields = [...(inherited ? [] : FRAME_FIELDS), ...OPTIONAL_FIELDS];
	for (const field of fields) {
		const value = record[field];
		if ((OPTIONAL_FIELDS as readonly string[]).includes(field) && value === undefined) continue;
		if (
			typeof value !== 'number' ||
			!Number.isFinite(value) ||
			!Number.isFinite(value * emuPerPx)
		) {
			throw new Error(
				`${context}: element '${element.id}' (${element.name ?? 'unnamed'}, type=${element.type}) has invalid ${field}=${String(value)}; expected a finite number`,
			);
		}
	}
	if (element.type === 'group') {
		for (const child of element.children) validateElementGeometry(child, context, emuPerPx);
	}
}

/** One preflight covers all shape, graphic-frame, media, ink, and group writers. */
export function validateSlideGeometry(slides: readonly PptxSlide[]): void {
	for (const slide of slides) {
		for (const element of slide.elements) {
			validateElementGeometry(element, `Slide ${slide.slideNumber} (${slide.id})`);
		}
	}
}
