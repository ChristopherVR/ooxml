import { XMLBuilder } from 'fast-xml-parser';

import { orderedXmlKey, stripXmlOrderMarkers } from '../../geometry/custom-geometry-command-order';
import type { PptxSmartArtDrawingShape, TextSegment, TextStyle, XmlObject } from '../../types';
import { applyTextBodySp3d } from '../../utils/text-body-sp3d-writer';
import { at } from '../../../../geometry/indexed';

const builder = new XMLBuilder({
	ignoreAttributes: false,
	attributeNamePrefix: '@_',
	suppressEmptyNode: true,
	format: false,
});

function normalizeHex(color: string | undefined): string | undefined {
	const hex = color?.replace(/^#/u, '').trim();
	return hex && /^[0-9A-Fa-f]{6}$/u.test(hex) ? hex.toUpperCase() : undefined;
}

export type SmartArtRunWriter = (style: TextStyle) => XmlObject | undefined;

function runProperties(style: TextStyle, writeRun?: SmartArtRunWriter): XmlObject {
	// WordArt picks replace the raw run properties and set the fill choice explicitly.
	const authored =
		!style.runPropertiesXml && style.textFillNone !== undefined ? writeRun?.(style) : undefined;
	if (authored) return authored;
	const rPr: XmlObject = style.runPropertiesXml
		? (JSON.parse(JSON.stringify(style.runPropertiesXml)) as XmlObject)
		: { '@_lang': style.language ?? 'en-US' };
	if (style.fontSize !== undefined) {
		rPr['@_sz'] = String(Math.round(style.fontSize * (72 / 96) * 100));
	}
	if (style.bold !== undefined) {
		rPr['@_b'] = style.bold ? '1' : '0';
	}
	if (style.italic !== undefined) {
		rPr['@_i'] = style.italic ? '1' : '0';
	}
	if (style.underlineStyle) {
		rPr['@_u'] = style.underlineStyle;
	} else if (style.underline) {
		rPr['@_u'] = 'sng';
	}
	const color = normalizeHex(style.color);
	if (style.colorXml) {
		rPr['a:solidFill'] = JSON.parse(JSON.stringify(style.colorXml)) as XmlObject;
	} else if (color) {
		rPr['a:solidFill'] = { 'a:srgbClr': { '@_val': color } };
	}
	if (style.fontFamily && !rPr['a:latin']) {
		rPr['a:latin'] = { '@_typeface': style.fontFamily };
	}
	if (style.runPropertiesExtLstXml) {
		rPr['a:extLst'] = style.runPropertiesExtLstXml;
	}
	return rPr;
}

function paragraphProperties(segment: TextSegment | undefined): XmlObject {
	const style = segment?.style ?? {};
	const pPr: XmlObject = {};
	const alignMap: Partial<Record<NonNullable<TextStyle['align']>, string>> = {
		left: 'l',
		center: 'ctr',
		right: 'r',
		justify: 'just',
		justLow: 'justLow',
		dist: 'dist',
		thaiDist: 'thaiDist',
	};
	const align = style.align ? alignMap[style.align] : undefined;
	if (align) {
		pPr['@_algn'] = align;
	}
	if (segment?.paragraphLevel) {
		pPr['@_lvl'] = String(segment.paragraphLevel);
	}
	const bullet = segment?.bulletInfo;
	if (bullet?.none) {
		pPr['a:buNone'] = {};
	} else if (bullet?.char) {
		pPr['a:buChar'] = { '@_char': bullet.char };
	} else if (bullet?.autoNumType) {
		pPr['a:buAutoNum'] = {
			'@_type': bullet.autoNumType,
			...(bullet.autoNumStartAt ? { '@_startAt': String(bullet.autoNumStartAt) } : {}),
		};
	}
	if (style.paragraphPropertiesExtLstXml) {
		pPr['a:extLst'] = style.paragraphPropertiesExtLstXml;
	}
	if (style.paragraphDefaultRunPropertiesXml) {
		pPr['a:defRPr'] = style.paragraphDefaultRunPropertiesXml;
	}
	return pPr;
}

function append(target: XmlObject, name: string, value: XmlObject, order: number): void {
	const repeated = Object.keys(target).some((key) => key === name || key.startsWith(`${name}#`));
	target[repeated ? orderedXmlKey(name, order) : name] = value;
}

function paragraphXml(segments: TextSegment[], writeRun?: SmartArtRunWriter): XmlObject {
	const paragraph: XmlObject = { 'a:pPr': paragraphProperties(segments[0]) };
	let order = 0;
	for (const segment of segments) {
		if (segment.isLineBreak) {
			append(
				paragraph,
				'a:br',
				{
					'a:rPr':
						(segment.breakRunProperties as XmlObject | undefined) ??
						runProperties(segment.style, writeRun),
				},
				order++,
			);
		} else if (segment.text === '\t') {
			append(paragraph, 'a:tab', {}, order++);
		} else if (segment.fieldType) {
			append(
				paragraph,
				'a:fld',
				{
					...(segment.fieldGuid ? { '@_id': segment.fieldGuid } : {}),
					'@_type': segment.fieldType,
					'a:rPr': runProperties(segment.style, writeRun),
					...(segment.fieldParagraphPropertiesXml
						? { 'a:pPr': segment.fieldParagraphPropertiesXml as XmlObject }
						: {}),
					'a:t': segment.text,
				},
				order++,
			);
		} else {
			append(
				paragraph,
				'a:r',
				{
					'a:rPr': runProperties(segment.style, writeRun),
					'a:t': segment.text,
				},
				order++,
			);
		}
	}
	const endProperties = segments[0]?.endParaRunProperties;
	if (endProperties) {
		paragraph['a:endParaRPr'] = endProperties as XmlObject;
	}
	return paragraph;
}

function splitParagraphs(segments: TextSegment[]): TextSegment[][] {
	const paragraphs: TextSegment[][] = [[]];
	for (const segment of segments) {
		if (segment.isParagraphBreak) {
			paragraphs.push([]);
		} else {
			at(paragraphs, paragraphs.length - 1).push(segment);
		}
	}
	return paragraphs;
}

function segmentText(segments: TextSegment[]): string {
	return segments.map((segment) => (segment.isParagraphBreak ? '\n' : segment.text)).join('');
}

function reconcileFlatText(segments: TextSegment[], desired: string): TextSegment[] {
	if (segmentText(segments) === desired) {
		return segments;
	}
	const separators = (value: string): string =>
		[...value].filter((character) => character === '\n' || character === '\t').join('');
	if (separators(segmentText(segments)) !== separators(desired)) {
		// With no segment to inherit from, the flattened text takes the default (empty) style.
		const [first = { text: '', style: {} }] = segments;
		return [{ ...first, text: desired, isLineBreak: undefined, isParagraphBreak: undefined }];
	}
	const result = segments.map((segment) => ({ ...segment }));
	const targets = desired.split(/[\t\n]/u);
	let targetIndex = 0;
	let region: number[] = [];
	const flush = (): void => {
		const target = targets[targetIndex++] ?? '';
		if (region.length > 0) {
			const texts = region.map((index) => at(result, index).text);
			const before = texts.join('');
			let prefix = 0;
			while (
				prefix < before.length &&
				prefix < target.length &&
				before[prefix] === target[prefix]
			) {
				prefix++;
			}
			let suffix = 0;
			while (
				suffix < before.length - prefix &&
				suffix < target.length - prefix &&
				before[before.length - 1 - suffix] === target[target.length - 1 - suffix]
			) {
				suffix++;
			}
			const removeStart = prefix;
			const removeEnd = before.length - suffix;
			const insertion = target.slice(prefix, target.length - suffix);
			let offset = 0;
			let inserted = false;
			for (let regionIndex = 0; regionIndex < region.length; regionIndex++) {
				// `texts` has one entry per region index.
				const text = at(texts, regionIndex);
				const start = offset;
				const end = start + text.length;
				const localStart = Math.max(0, removeStart - start);
				const localEnd = Math.min(text.length, removeEnd - start);
				if (localEnd > localStart || (!inserted && removeStart >= start && removeStart <= end)) {
					at(result, at(region, regionIndex)).text =
						text.slice(0, localStart) + (inserted ? '' : insertion) + text.slice(localEnd);
					inserted = true;
				}
				offset = end;
			}
		}
		region = [];
	};
	for (let index = 0; index < result.length; index++) {
		const segment = at(result, index);
		if (segment.isParagraphBreak || segment.isLineBreak || segment.text === '\t') {
			flush();
		} else {
			region.push(index);
		}
	}
	flush();
	return result;
}

/** Build a cached `dsp:txBody` from standard shape text segments. */
export function drawingTextBodyXml(
	shape: PptxSmartArtDrawingShape,
	writeRun?: SmartArtRunWriter,
): string {
	const fallbackStyle: TextStyle = {
		fontSize: shape.fontSize,
		color: shape.fontColor,
		align: 'center',
	};
	const sourceSegments = shape.textSegments?.length
		? shape.textSegments
		: [{ text: shape.text ?? '', style: fallbackStyle }];
	const segments = reconcileFlatText(sourceSegments, shape.text ?? segmentText(sourceSegments));
	const paragraphs = splitParagraphs(segments).map((items) =>
		paragraphXml(items.length > 0 ? items : [{ text: '', style: fallbackStyle }], writeRun),
	);
	const bodyPr: XmlObject = { '@_anchor': 'ctr' };
	applyTextBodySp3d(bodyPr, { text3d: shape.text3d });
	return stripXmlOrderMarkers(
		builder.build({
			'dsp:txBody': {
				'a:bodyPr': bodyPr,
				'a:lstStyle': {},
				'a:p': paragraphs.length === 1 ? paragraphs[0] : paragraphs,
			},
		}),
	);
}
