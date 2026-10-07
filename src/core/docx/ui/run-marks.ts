import type { TextRun } from '../model.js';
import type { Mark, Schema } from 'prosemirror-model';
import { expectDefined } from './defined.js';
import { extraRunProperties } from './run-extra-mark.js';

/** ProseMirror marks carrying a text run's character formatting, links, fields and revisions. */
export function marksForRun(run: TextRun, schema: Schema): Mark[] {
	const marks: Mark[] = [];
	if (run.link)
		marks.push(
			expectDefined(schema.marks.link, 'link mark').create({
				href: run.link.href ?? null,
				anchor: run.link.anchor ?? null,
				tooltip: run.link.tooltip ?? null,
			}),
		);
	if (run.bold) marks.push(expectDefined(schema.marks.bold, 'bold mark').create());
	if (run.italic) marks.push(expectDefined(schema.marks.italic, 'italic mark').create());
	if (run.underline) marks.push(expectDefined(schema.marks.underline, 'underline mark').create());
	if (run.strike) marks.push(expectDefined(schema.marks.strike, 'strike mark').create());
	if (run.highlight)
		marks.push(
			expectDefined(schema.marks.highlight, 'highlight mark').create({ color: run.highlight }),
		);
	if (run.verticalAlign)
		marks.push(
			expectDefined(schema.marks.verticalAlign, 'verticalAlign mark').create({
				value: run.verticalAlign,
			}),
		);
	if (
		run.language !== undefined ||
		run.eastAsiaLanguage !== undefined ||
		run.bidiLanguage !== undefined
	)
		marks.push(
			expectDefined(schema.marks.language, 'language mark').create({
				language: run.language ?? null,
				eastAsiaLanguage: run.eastAsiaLanguage ?? null,
				bidiLanguage: run.bidiLanguage ?? null,
			}),
		);
	if (run.rtl !== undefined)
		marks.push(expectDefined(schema.marks.runRtl, 'runRtl mark').create({ value: run.rtl }));
	if (run.fontFamily || run.fontSize || run.color)
		marks.push(
			expectDefined(schema.marks.font, 'font mark').create({
				family: run.fontFamily || null,
				size: run.fontSize || null,
				color: run.color || null,
			}),
		);
	if (run.field)
		marks.push(
			expectDefined(schema.marks.field, 'field mark').create({
				instr: run.field.instr,
				simple: Boolean(run.field.simple),
			}),
		);
	const revision = run.revision;
	const move = revision?.move ? JSON.stringify(revision.move) : null;
	if (revision?.kind === 'insert' || revision?.kind === 'moveTo')
		marks.push(
			expectDefined(schema.marks.insertion, 'insertion mark').create({
				author: revision.author,
				date: revision.date ?? null,
				dateUtc: revision.dateUtc ?? null,
				id: revision.id,
				move: revision.kind === 'moveTo' ? move : null,
			}),
		);
	else if (revision?.kind === 'delete' || revision?.kind === 'moveFrom')
		marks.push(
			expectDefined(schema.marks.deletion, 'deletion mark').create({
				author: revision.author,
				date: revision.date ?? null,
				dateUtc: revision.dateUtc ?? null,
				id: revision.id,
				move: revision.kind === 'moveFrom' ? move : null,
			}),
		);
	for (const id of [...new Set(run.commentIds ?? [])].sort())
		marks.push(expectDefined(schema.marks.comment, 'comment mark').create({ ids: [id] }));
	if (run.style)
		marks.push(
			expectDefined(schema.marks.characterStyle, 'characterStyle mark').create({ id: run.style }),
		);
	const extra = extraRunProperties(run);
	if (extra)
		marks.push(
			expectDefined(schema.marks.runProperties, 'runProperties mark').create({ props: extra }),
		);
	return marks;
}
