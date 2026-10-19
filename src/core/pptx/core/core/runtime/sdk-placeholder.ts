import type { PptxElement, XmlObject } from '../../types';
import { scorePlaceholderMatch } from '../../utils/placeholder-remap';
import { setRawXmlTransform } from '../../utils/placeholder-xml';
import { canonicalPlaceholderType } from '../../utils/placeholder-validation';
import { xmlPath } from '../../utils/xml-access';

/** Text-capable placeholder families; other slots require non-text content. */
const TEXT_PLACEHOLDER_TYPES = new Set([
	'title',
	'ctrtitle',
	'subtitle',
	'body',
	'obj',
	'dt',
	'ftr',
	'hdr',
	'sldnum',
]);

export interface SdkPlaceholderOptions {
	idx?: number;
	x?: number;
	y?: number;
	width?: number;
	height?: number;
}

/** Pick a text-capable slot using the layout switcher's existing role matcher. */
export function selectTextPlaceholder(
	input: PptxElement,
	type: string,
	options: SdkPlaceholderOptions,
	slots: Array<{ phInfo: { type?: string; idx?: string }; shapeXml: XmlObject }>,
): { type: string; idx?: string } {
	if (!TEXT_PLACEHOLDER_TYPES.has(type.toLowerCase())) {
		throw new Error(`Unsupported text placeholder type: ${type}`);
	}
	if (
		options.idx !== undefined &&
		(!Number.isInteger(options.idx) || options.idx < 0 || options.idx > 4294967295)
	) {
		throw new Error('Placeholder idx must be an unsigned 32-bit integer');
	}
	const source = {
		type: type.toLowerCase(),
		...(options.idx !== undefined ? { idx: String(options.idx) } : {}),
	};
	let best: (typeof slots)[number] | undefined;
	let bestScore = 0;
	for (const slot of slots) {
		if (!TEXT_PLACEHOLDER_TYPES.has(slot.phInfo.type ?? 'obj')) continue;
		if (source.idx !== undefined && source.idx !== (slot.phInfo.idx ?? '0')) continue;
		const score = scorePlaceholderMatch(input, source, slot.phInfo);
		if (score > bestScore) {
			best = slot;
			bestScore = score;
		}
	}
	return best
		? { type: canonicalPlaceholderType(best.phInfo.type ?? 'obj')!, idx: best.phInfo.idx }
		: { type: canonicalPlaceholderType(source.type)!, idx: source.idx };
}

/**
 * SDK text input has no language, so the run writer falls back to
 * `lang="en-US"` (as PowerPoint does for freshly typed text). On a placeholder
 * that would override the language the layout and master cascade supplies, so
 * drop the fallback, and any property element it leaves holding only the
 * `dirty="0"` marker. Every other property the caller set is kept.
 */
function withoutFallbackLanguage(properties: XmlObject): XmlObject | undefined {
	const { '@_lang': language, ...rest } = properties;
	if (language !== 'en-US') return properties;
	const keys = Object.keys(rest);
	return keys.length === 0 || (keys.length === 1 && rest['@_dirty'] === '0') ? undefined : rest;
}

function stripFallbackLanguage(node: XmlObject, key: 'a:rPr' | 'a:endParaRPr'): XmlObject {
	const properties = node[key];
	if (!properties || typeof properties !== 'object') return node;
	const next = withoutFallbackLanguage(properties as XmlObject);
	if (next === properties) return node;
	const result: XmlObject = {};
	for (const [name, value] of Object.entries(node)) {
		if (name !== key) result[name] = value;
		else if (next) result[name] = next;
	}
	return result;
}

function withoutFallbackLanguages(paragraphs: XmlObject[]): XmlObject[] {
	return paragraphs.map((paragraph) => {
		const next = { ...stripFallbackLanguage(paragraph, 'a:endParaRPr') };
		for (const child of ['a:r', 'a:br', 'a:fld'] as const) {
			const nodes = next[child];
			if (nodes === undefined) continue;
			next[child] = Array.isArray(nodes)
				? (nodes as XmlObject[]).map((node) => stripFallbackLanguage(node, 'a:rPr'))
				: stripFallbackLanguage(nodes as XmlObject, 'a:rPr');
		}
		return next;
	});
}

/** A real slide placeholder with no authored frame or text-format defaults. */
export function placeholderShape(
	identity: { type: string; idx?: string },
	paragraphs: XmlObject[],
): XmlObject {
	paragraphs = withoutFallbackLanguages(paragraphs);
	return {
		'p:nvSpPr': {
			'p:cNvPr': { '@_id': '2', '@_name': `Placeholder ${identity.type}` },
			'p:cNvSpPr': { 'a:spLocks': { '@_noGrp': '1' } },
			'p:nvPr': {
				'p:ph': {
					'@_type': identity.type,
					...(identity.idx !== undefined ? { '@_idx': identity.idx } : {}),
				},
			},
		},
		'p:spPr': {},
		'p:txBody': { 'a:bodyPr': {}, 'a:lstStyle': {}, 'a:p': paragraphs },
	};
}

/** Validate inherited geometry before parsing can degrade a missing frame to zeros. */
export function applyPlaceholderFrame(
	shape: XmlObject,
	inherited: XmlObject | undefined,
	options: SdkPlaceholderOptions,
	emuPerPx: number,
	materializeFrame = false,
): void {
	for (const key of ['x', 'y', 'width', 'height'] as const) {
		if (options[key] !== undefined && !Number.isFinite(options[key])) {
			throw new Error(`Invalid placeholder frame for ${key}: expected a finite number`);
		}
	}
	const off = xmlPath(inherited, 'p:spPr', 'a:xfrm', 'a:off');
	const ext = xmlPath(inherited, 'p:spPr', 'a:xfrm', 'a:ext');
	const frame = {
		x: options.x ?? (off?.['@_x'] !== undefined ? Number(off['@_x']) / emuPerPx : undefined),
		y: options.y ?? (off?.['@_y'] !== undefined ? Number(off['@_y']) / emuPerPx : undefined),
		width:
			options.width ?? (ext?.['@_cx'] !== undefined ? Number(ext['@_cx']) / emuPerPx : undefined),
		height:
			options.height ?? (ext?.['@_cy'] !== undefined ? Number(ext['@_cy']) / emuPerPx : undefined),
	};
	for (const [key, value] of Object.entries(frame)) {
		if (!Number.isFinite(value)) {
			throw new Error(
				`No matching placeholder frame for ${key}; supply a complete explicit frame or choose a layout containing this slot`,
			);
		}
	}
	if (
		materializeFrame ||
		(['x', 'y', 'width', 'height'] as const).some((key) => options[key] !== undefined)
	) {
		setRawXmlTransform(
			shape,
			...((['x', 'y', 'width', 'height'] as const).map((key) =>
				Math.round(frame[key]! * emuPerPx),
			) as [number, number, number, number]),
		);
	}
}
