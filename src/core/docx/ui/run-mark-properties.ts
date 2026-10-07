import type { TextRun } from '../model.js';
import type { Node as ProseMirrorNode, Mark } from 'prosemirror-model';
import { explicitOffFields } from './run-extra-mark.js';
import { commentIdsFromMarks } from './comment-anchors.js';

function propertyOfMark(child: ProseMirrorNode, name: string): Mark | undefined {
	return child.marks.find((mark) => mark.type.name === name);
}

export function linkFromMarks(child: ProseMirrorNode): TextRun['link'] {
	const link = propertyOfMark(child, 'link');
	if (!link) return undefined;
	const info: NonNullable<TextRun['link']> = {};
	if (link.attrs.href) info.href = link.attrs.href;
	if (link.attrs.anchor) info.anchor = link.attrs.anchor;
	if (link.attrs.tooltip) info.tooltip = link.attrs.tooltip;
	return info.href || info.anchor ? info : undefined;
}

/** Copies run formatting carried by a node's marks onto `run` (text runs and note references). */
export function applyMarkFormatting(run: TextRun, child: ProseMirrorNode): void {
	if (propertyOfMark(child, 'bold')) run.bold = true;
	if (propertyOfMark(child, 'italic')) run.italic = true;
	if (propertyOfMark(child, 'underline')) run.underline = true;
	if (propertyOfMark(child, 'strike')) run.strike = true;
	const highlight = propertyOfMark(child, 'highlight');
	if (highlight?.attrs.color) run.highlight = highlight.attrs.color;
	const verticalAlign = propertyOfMark(child, 'verticalAlign');
	if (verticalAlign?.attrs.value) run.verticalAlign = verticalAlign.attrs.value;
	const language = propertyOfMark(child, 'language');
	if (language?.attrs.language != null) run.language = language.attrs.language;
	if (language?.attrs.eastAsiaLanguage != null)
		run.eastAsiaLanguage = language.attrs.eastAsiaLanguage;
	if (language?.attrs.bidiLanguage != null) run.bidiLanguage = language.attrs.bidiLanguage;
	const rtl = propertyOfMark(child, 'runRtl');
	if (rtl) run.rtl = rtl.attrs.value;
	const font = propertyOfMark(child, 'font');
	if (font?.attrs.family) run.fontFamily = font.attrs.family;
	if (font?.attrs.size) run.fontSize = font.attrs.size;
	if (font?.attrs.color) run.color = font.attrs.color;
	const insertion = propertyOfMark(child, 'insertion');
	const deletion = propertyOfMark(child, 'deletion');
	const revisionMark = insertion ?? deletion;
	if (revisionMark) {
		const move =
			typeof revisionMark.attrs.move === 'string'
				? (JSON.parse(revisionMark.attrs.move) as NonNullable<TextRun['revision']>['move'])
				: undefined;
		run.revision = {
			// A move without its move data is recorded as a plain insertion or deletion.
			kind: insertion ? (move ? 'moveTo' : 'insert') : move ? 'moveFrom' : 'delete',
			author: String(revisionMark.attrs.author || ''),
			id: String(revisionMark.attrs.id || ''),
			...(revisionMark.attrs.date ? { date: String(revisionMark.attrs.date) } : {}),
			...(revisionMark.attrs.dateUtc ? { dateUtc: String(revisionMark.attrs.dateUtc) } : {}),
			...(move ? { move } : {}),
		};
	}
	const commentIds = commentIdsFromMarks(child.marks);
	if (commentIds.length) run.commentIds = commentIds;
	const field = propertyOfMark(child, 'field');
	if (field)
		run.field = {
			instr: String(field.attrs.instr),
			...(field.attrs.simple ? { simple: true } : {}),
		};
	const characterStyle = propertyOfMark(child, 'characterStyle');
	if (characterStyle?.attrs.id) run.style = String(characterStyle.attrs.id);
	const extra = propertyOfMark(child, 'runProperties');
	if (extra?.attrs.props)
		for (const [key, value] of Object.entries(structuredClone(extra.attrs.props) as TextRun)) {
			// An explicit off never overrides a mark the user applied (bold on text that was unbolded).
			const field = key as keyof TextRun;
			if (key === 'revision' && run.revision) {
				if ((value as TextRun['revision'])?.kind === 'formatChange')
					run.formatRevision = value as NonNullable<TextRun['revision']>;
				continue;
			}
			if ((explicitOffFields as readonly string[]).includes(key) && run[field] !== undefined)
				continue;
			setRunField(run, field, value);
		}
}

/** Typed per-field assignment, so a value can only be written to a field that accepts it. */
function setRunField<K extends keyof TextRun>(run: TextRun, field: K, value: TextRun[K]): void {
	run[field] = value;
}
