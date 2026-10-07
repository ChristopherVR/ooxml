import type { TextRun } from '../model';
import type { Node as ProseMirrorNode, Schema } from 'prosemirror-model';

import { extraRunProperties } from './run-extra-mark';
import { applyMarkFormatting, linkFromMarks } from './run-mark-properties';
import { marksForRun } from './run-marks';

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
		const { text: _text, image: _image, link: _link, ...format } = run;
		const marks = run.link
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
					anchored: Boolean(run.image.anchored),
					placement: run.image.placement ? JSON.stringify(run.image.placement) : null,
					svgPartName: run.image.svgPartName ?? null,
					unsupported: run.image.unsupported ?? null,
					textBoxText: run.image.textBoxText ? JSON.stringify(run.image.textBoxText) : null,
					textBoxEditable: run.image.textBoxEditable ?? null,
					textBoxBorder: run.image.textBoxBorder ?? null,
					diagram: run.image.diagram ? JSON.stringify(run.image.diagram) : null,
				},
				undefined,
				marks,
			),
		];
	}
	if (!run.text) return [];
	const marks = marksForRun(run, schema);
	return run.text
		.split(/(\n)/)
		.flatMap((piece) =>
			piece === '\n'
				? [schema.node('hardBreak', null, undefined, marks)]
				: piece
					? [schema.text(piece, marks)]
					: [],
		);
}

function applyInlineFormat(run: TextRun, child: ProseMirrorNode): void {
	if (typeof child.attrs.format === 'string')
		Object.assign(run, JSON.parse(child.attrs.format) as Partial<TextRun>);
	// Imported atom attributes are the base; newly applied marks may explicitly turn them off.
	// Keep text revision context when an old formatting revision rides in the extra-properties mark.
	const marks: TextRun = { text: '', ...(run.revision ? { revision: run.revision } : {}) };
	applyMarkFormatting(marks, child);
	const { text: _text, ...properties } = marks;
	Object.assign(run, properties);
}

/** A single source run, including inline atoms with their own properties and revisions. */
export function inlineNodeRun(child: ProseMirrorNode): TextRun | undefined {
	const runs: TextRun[] = [];
	appendInlineNode(runs, child);
	return runs[0];
}

export function appendInlineNode(runs: TextRun[], child: ProseMirrorNode): void {
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
				...(child.attrs.anchored ? { anchored: true } : {}),
				...(child.attrs.svgPartName ? { svgPartName: child.attrs.svgPartName } : {}),
				// Same key order as the parser's run: the writer compares an unchanged placeholder run by JSON.
				...(child.attrs.unsupported ? { unsupported: child.attrs.unsupported } : {}),
				...(typeof child.attrs.diagram === 'string'
					? { diagram: JSON.parse(child.attrs.diagram) }
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
	applyMarkFormatting(run, child);
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
		previous &&
		!previous.break &&
		!previous.image &&
		!previous.equation &&
		!previous.noteReference &&
		!previous.fieldChar &&
		previous.fieldCode === undefined &&
		JSON.stringify(previous.field) === JSON.stringify(run.field) &&
		sameRevision &&
		sameComments &&
		sameExtra &&
		sameLink &&
		fields.every((field) => previous[field] === run[field])
	)
		previous.text += run.text;
	else runs.push(run);
}
