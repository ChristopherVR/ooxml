/**
 * Word writes a field's code inside `<!--[if supportFields]>` comments, so browsers show only
 * its result: `field-begin`, the code, `field-separator`, then the visible result and a
 * `field-end`. The tracker reads those comments and wraps the outermost result in the editor's
 * field span (`data-field`), so a pasted PAGE or DATE field stays a field.
 */
import { parseDeclarations } from './css';
import { decodeEntities, escapeAttribute, tokenizeHtml } from './html-tokens';

interface Field {
	code: string;
	phase: 'code' | 'result';
	/** The element depth its result span was opened at, while open. */
	openAt?: number | undefined;
}

export class FieldTracker {
	private readonly fields: Field[] = [];

	/** Whether text belongs to a field code rather than to the document. */
	get inCode(): boolean {
		return this.fields.some((field) => field.phase === 'code');
	}

	addCode(text: string): void {
		const field = this.fields.find((candidate) => candidate.phase === 'code');
		if (field?.phase === 'code') field.code += text;
	}

	/** Reads one comment at element depth `depth`; returns the markup it opens or closes. */
	comment(text: string, depth: number): string {
		const body = /^\[if\s+supportFields\]>([\s\S]*?)<!\[endif\]$/i.exec(text.trim())?.[1];
		if (body === undefined) return '';
		let out = '';
		for (const token of tokenizeHtml(body)) {
			if (token.kind === 'text') {
				this.addCode(decodeEntities(token.text));
				continue;
			}
			if (token.kind !== 'open' || token.name !== 'span') continue;
			const element = parseDeclarations(token.attrs.get('style')).get('mso-element');
			if (element === 'field-begin') this.fields.push({ code: '', phase: 'code' });
			else if (element === 'field-separator') out += this.separate(depth);
			else if (element === 'field-end') out += this.end(depth);
		}
		return out;
	}

	private separate(depth: number): string {
		const field = this.fields.at(-1);
		if (!field || field.phase !== 'code') return '';
		field.phase = 'result';
		const code = field.code.replace(/\s+/g, ' ').trim();
		// Only the outermost field's result is wrapped; nested results are part of its text.
		if (this.fields.length !== 1 || !code) return '';
		field.openAt = depth;
		return `<span data-field="${escapeAttribute(` ${code} `)}" data-field-simple="1">`;
	}

	private end(depth: number): string {
		const field = this.fields.pop();
		if (field?.openAt === undefined || field.openAt !== depth) return '';
		return '</span>';
	}

	/** Closes an open result span before an element it was opened inside closes (to `depth`). */
	closeTo(depth: number): string {
		const field = this.fields[0];
		if (field?.openAt === undefined || depth >= field.openAt) return '';
		field.openAt = undefined;
		return '</span>';
	}
}
