import type { TextRun } from '../model';
import type { DocxInk } from '../ink';
import type { Node as ProseMirrorNode, Schema } from 'prosemirror-model';

import { extraRunProperties } from './run-extra-mark';
import { applyMarkFormatting, linkFromMarks } from './run-mark-properties';
import { marksForRun } from './run-marks';
import { inlineRunProperties } from './inline-run-properties';
import { commentIdsFromNode } from './comment-anchors';

/** A note reference's number in document order and its label in the document's number format. */
export type NoteNumberLookup = (
	kind: 'footnote' | 'endnote',
	id: string,
) => { number: number; label: string };

export function runToInlineNodes(
	run: TextRun,
	schema: Schema,
	noteNumber?: NoteNumberLookup,
): ProseMirrorNode[] {
	if (run.equation) {
		const { text: _text, equation, ...format } = run;
		return [
			schema.node('equation', {
				omml: equation.omml,
				display: equation.display,
				format: Object.keys(format).length ? JSON.stringify(format) : null,
			}),
		];
	}
	if (run.break) {
		const { text: _text, break: kind, ...format } = run;
		return [
			schema.node('pageBreak', {
				kind,
				format: Object.keys(format).length ? JSON.stringify(format) : null,
			}),
		];
	}
	if (run.fieldChar || run.fieldCode !== undefined) {
		const { text: _text, fieldChar, fieldCode, ...format } = run;
		return [
			schema.node('fieldMarker', {
				kind: fieldChar ?? 'code',
				code: fieldCode ?? null,
				format: Object.keys(format).length ? JSON.stringify(format) : null,
			}),
		];
	}
	if (run.noteReference) {
		const { kind, id } = run.noteReference;
		const { text: _text, noteReference: _reference, ...format } = run;
		return [
			schema.node('noteReference', {
				kind,
				id,
				number: noteNumber?.(kind, id).number ?? 1,
				label: noteNumber?.(kind, id).label ?? null,
				format: Object.keys(format).length ? JSON.stringify(format) : null,
			}),
		];
	}
	if (run.image) {
		const { text: _text, image: _image, ...format } = run;
		const marks =
			run.link && !schema.nodes.image?.spec.attrs?.runFormat_bold
				? [
						schema.mark('link', {
							href: run.link.href ?? null,
							anchor: run.link.anchor ?? null,
							tooltip: run.link.tooltip ?? null,
						}),
					]
				: [];
		return [
			schema.node(
				'image',
				{
					format: Object.keys(format).length ? JSON.stringify(format) : null,
					relId: run.image.relId,
					partName: run.image.partName,
					contentType: run.image.contentType,
					widthPx: run.image.widthPx,
					heightPx: run.image.heightPx,
					altText: run.image.altText ?? null,
					title: run.image.title ?? null,
					drawingId: run.image.drawingId ?? null,
					drawingName: run.image.drawingName ?? null,
					anchored: Boolean(run.image.anchored),
					placement: run.image.placement ? JSON.stringify(run.image.placement) : null,
					svgPartName: run.image.svgPartName ?? null,
					transform: run.image.transform ? JSON.stringify(run.image.transform) : null,
					pictureFill: run.image.pictureFill ? JSON.stringify(run.image.pictureFill) : null,
					vmlFloat: run.image.vmlFloat === true,
					unsupported: run.image.unsupported ?? null,
					textBoxText: run.image.textBoxText ? JSON.stringify(run.image.textBoxText) : null,
					textBoxEditable: run.image.textBoxEditable ?? null,
					textBoxBorder: run.image.textBoxBorder ?? null,
					textBoxContent: run.image.textBoxContent
						? JSON.stringify(run.image.textBoxContent)
						: null,
					diagram: run.image.diagram ? JSON.stringify(run.image.diagram) : null,
					chart: run.image.chart ? JSON.stringify(run.image.chart) : null,
					payload: run.image.payload ? JSON.stringify(run.image.payload) : null,
					shape: run.image.shape ? JSON.stringify(run.image.shape) : null,
					ink: run.image.ink ? JSON.stringify(run.image.ink) : null,
					legacyInk: run.image.legacyInk ? JSON.stringify(run.image.legacyInk) : null,
					watermark: run.image.watermark ? JSON.stringify(run.image.watermark) : null,
					pictureWatermark: run.image.pictureWatermark
						? JSON.stringify(run.image.pictureWatermark)
						: null,
					signatureLine: run.image.signatureLine ? JSON.stringify(run.image.signatureLine) : null,
					webVideo: run.image.webVideo ? JSON.stringify(run.image.webVideo) : null,
				},
				undefined,
				marks,
			),
		];
	}
	if (!run.text) {
		const point = schema.nodes.commentPoint;
		return point ? (run.commentIds ?? []).map((id) => point.create({ id })) : [];
	}
	const marks = marksForRun(run, schema);
	const { text: _text, ...format } = run;
	const hardBreak = schema.nodes.hardBreak;
	const breakNode = () =>
		hardBreak?.spec.attrs?.format
			? schema.node('hardBreak', {
					format: Object.keys(format).length ? JSON.stringify(format) : null,
				})
			: schema.node('hardBreak', null, undefined, marks);
	return run.text
		.split(/(\n)/)
		.flatMap((piece) =>
			piece === '\n' ? [breakNode()] : piece ? [schema.text(piece, marks)] : [],
		);
}

function applyInlineFormat(run: TextRun, child: ProseMirrorNode): void {
	Object.assign(run, inlineRunProperties(child));
	// Imported atom attributes are the base; newly applied marks may explicitly turn them off.
	// Keep text revision context when an old formatting revision rides in the extra-properties mark.
	const marks: TextRun = { text: '', ...(run.revision ? { revision: run.revision } : {}) };
	applyMarkFormatting(marks, child);
	const { text: _text, ...properties } = marks;
	Object.assign(run, properties);
	const ids = commentIdsFromNode(child);
	if (ids.length) run.commentIds = ids;
	else delete run.commentIds;
}

/** A single source run, including inline atoms with their own properties and revisions. */
export function inlineNodeRun(child: ProseMirrorNode): TextRun | undefined {
	const runs: TextRun[] = [];
	appendInlineNode(runs, child);
	return runs[0];
}

export function appendInlineNode(runs: TextRun[], child: ProseMirrorNode): void {
	if (child.type.name === 'commentPoint') {
		const commentIds = commentIdsFromNode(child);
		if (commentIds.length) runs.push({ text: '', commentIds });
		return;
	}
	if (child.type.name === 'equation') {
		const run: TextRun = {
			text: '',
			equation: { omml: String(child.attrs.omml), display: Boolean(child.attrs.display) },
		};
		applyInlineFormat(run, child);
		runs.push(run);
		return;
	}
	if (child.type.name === 'pageBreak') {
		const run: TextRun = { text: '', break: child.attrs.kind === 'column' ? 'column' : 'page' };
		applyInlineFormat(run, child);
		runs.push(run);
		return;
	}
	if (child.type.name === 'fieldMarker') {
		const marker: TextRun =
			child.attrs.kind === 'code'
				? { text: '', fieldCode: String(child.attrs.code ?? '') }
				: { text: '', fieldChar: child.attrs.kind };
		applyInlineFormat(marker, child);
		runs.push(marker);
		return;
	}
	if (child.type.name === 'noteReference') {
		const reference: TextRun = {
			text: '',
			noteReference: {
				kind: child.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
				id: String(child.attrs.id || ''),
			},
		};
		applyInlineFormat(reference, child);
		runs.push(reference);
		return;
	}
	if (child.type.name === 'image') {
		const imageLink = linkFromMarks(child);
		runs.push({
			text: '',
			image: {
				relId: child.attrs.relId,
				partName: child.attrs.partName,
				contentType: child.attrs.contentType,
				widthPx: child.attrs.widthPx,
				heightPx: child.attrs.heightPx,
				...(child.attrs.altText ? { altText: child.attrs.altText } : {}),
				...(child.attrs.title ? { title: child.attrs.title } : {}),
				...(child.attrs.drawingId ? { drawingId: child.attrs.drawingId } : {}),
				...(child.attrs.drawingName ? { drawingName: child.attrs.drawingName } : {}),
				...(child.attrs.anchored ? { anchored: true } : {}),
				...(child.attrs.vmlFloat === true ? { vmlFloat: true } : {}),
				...(child.attrs.svgPartName ? { svgPartName: child.attrs.svgPartName } : {}),
				...(typeof child.attrs.transform === 'string'
					? { transform: JSON.parse(child.attrs.transform) }
					: {}),
				...(typeof child.attrs.pictureFill === 'string'
					? { pictureFill: JSON.parse(child.attrs.pictureFill) }
					: {}),
				// Same key order as the parser's run: the writer compares an unchanged placeholder run by JSON.
				...(child.attrs.unsupported ? { unsupported: child.attrs.unsupported } : {}),
				...(typeof child.attrs.diagram === 'string'
					? { diagram: JSON.parse(child.attrs.diagram) }
					: {}),
				...(typeof child.attrs.chart === 'string' ? { chart: JSON.parse(child.attrs.chart) } : {}),
				...(typeof child.attrs.payload === 'string'
					? { payload: JSON.parse(child.attrs.payload) }
					: {}),
				...(typeof child.attrs.shape === 'string' ? { shape: JSON.parse(child.attrs.shape) } : {}),
				...(typeof child.attrs.ink === 'string' ? { ink: inkAttr(child.attrs.ink) } : {}),
				...(typeof child.attrs.legacyInk === 'string'
					? { legacyInk: JSON.parse(child.attrs.legacyInk) }
					: {}),
				...(typeof child.attrs.watermark === 'string'
					? { watermark: JSON.parse(child.attrs.watermark) }
					: {}),
				...(typeof child.attrs.pictureWatermark === 'string'
					? { pictureWatermark: JSON.parse(child.attrs.pictureWatermark) }
					: {}),
				...(typeof child.attrs.signatureLine === 'string'
					? { signatureLine: JSON.parse(child.attrs.signatureLine) }
					: {}),
				...(typeof child.attrs.webVideo === 'string'
					? { webVideo: JSON.parse(child.attrs.webVideo) }
					: {}),
				...(typeof child.attrs.placement === 'string'
					? { placement: JSON.parse(child.attrs.placement) }
					: {}),
				...(typeof child.attrs.textBoxText === 'string'
					? { textBoxText: JSON.parse(child.attrs.textBoxText) }
					: {}),
				...(child.attrs.textBoxEditable ? { textBoxEditable: true } : {}),
				...(child.attrs.textBoxEditable && child.attrs.textBoxBorder !== null
					? { textBoxBorder: Boolean(child.attrs.textBoxBorder) }
					: {}),
				...(typeof child.attrs.textBoxContent === 'string'
					? { textBoxContent: JSON.parse(child.attrs.textBoxContent) }
					: {}),
			},
			...(imageLink && { link: imageLink }),
		});
		applyInlineFormat(runs.at(-1)!, child);
		return;
	}
	if (!child.isText && child.type.name !== 'hardBreak') return;
	const run: TextRun = { text: child.isText ? child.text || '' : '\n' };
	const link = linkFromMarks(child);
	if (link) run.link = link;
	if (!child.isText && child.type.spec.attrs?.format) applyInlineFormat(run, child);
	else applyMarkFormatting(run, child);
	const previous = runs.at(-1);
	const fields: (keyof TextRun)[] = [
		'bold',
		'italic',
		'underline',
		'strike',
		'highlight',
		'verticalAlign',
		'language',
		'eastAsiaLanguage',
		'bidiLanguage',
		'rtl',
		'fontFamily',
		'fontSize',
		'color',
		'style',
	];
	const sameRevision = JSON.stringify(previous?.revision) === JSON.stringify(run.revision);
	const sameComments = JSON.stringify(previous?.commentIds) === JSON.stringify(run.commentIds);
	const sameLink = JSON.stringify(previous?.link ?? null) === JSON.stringify(run.link ?? null);
	const sameExtra =
		JSON.stringify(previous && extraRunProperties(previous)) ===
		JSON.stringify(extraRunProperties(run));
	if (
		previous?.text &&
		!previous.break &&
		!previous.image &&
		!previous.equation &&
		!previous.noteReference &&
		!previous.fieldChar &&
		previous.fieldCode === undefined &&
		JSON.stringify(previous.field) === JSON.stringify(run.field) &&
		JSON.stringify(previous.contentControl) === JSON.stringify(run.contentControl) &&
		sameRevision &&
		sameComments &&
		sameExtra &&
		sameLink &&
		fields.every((field) => previous[field] === run[field])
	)
		previous.text += run.text;
	else runs.push(run);
}

/**
 * Parsed ink attributes by text: the same object for unchanged ink, so the painters' caches keyed
 * by it hold across the edits of a session (a stroke drawn elsewhere re-reads every drawing).
 */
const inks = new Map<string, DocxInk>();
function inkAttr(text: string): DocxInk {
	const cached = inks.get(text);
	if (cached) return cached;
	if (inks.size >= 64) inks.delete(inks.keys().next().value as string);
	const ink = JSON.parse(text) as DocxInk;
	inks.set(text, ink);
	return ink;
}
