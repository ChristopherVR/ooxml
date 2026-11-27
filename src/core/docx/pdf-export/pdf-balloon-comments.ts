/**
 * The comment balloons of the Word PDF export's markup pane: each thread as Word labels it
 * ("Commented [AL1]: ...", replies "Commented [BB2R1]: ...", numbered in document order), with
 * the author and date under each entry, resolved comments greyed. DOM-free.
 */
import type { Comment } from '../index';
import type { BalloonPiece } from './pdf-balloons';

const RESOLVED = '#7f7f7f';
const META = '#595959';

/** A colour mixed with white, for comment highlights and balloon fills. */
export function paleColor(color: string, amount = 0.8): string {
	const hex = /^#?([\da-f]{6})$/i.exec(color)?.[1];
	if (!hex) return color;
	const mix = (index: number) =>
		Math.round(Number.parseInt(hex.slice(index, index + 2), 16) * (1 - amount) + 255 * amount)
			.toString(16)
			.padStart(2, '0');
	return `#${mix(0)}${mix(2)}${mix(4)}`;
}

/** The comment's initials, or the author's when the comment has none. */
export const initialsOf = (comment: Comment): string =>
	comment.initials ||
	comment.author
		.split(/\s+/)
		.filter(Boolean)
		.map((word) => word[0]!.toUpperCase())
		.join('');

/** `2026-03-04T05:06:07Z` as `2026-03-04 05:06`, as written (no time zone conversion). */
export const shortDate = (date: string | undefined): string =>
	date
		? (/^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2}))?/.exec(date)?.slice(1).filter(Boolean).join(' ') ??
			date)
		: '';

export class DocxPdfBalloonComments {
	readonly #comments = new Map<string, { comment: Comment; number: number }>();
	readonly #replies = new Map<string, Comment[]>();

	constructor(
		comments: readonly Comment[] | undefined,
		/** The document's font (comment text) and the label font. */
		readonly fonts: { text: string; label: string },
	) {
		(comments ?? []).forEach((comment, index) => {
			this.#comments.set(comment.id, { comment, number: index + 1 });
			if (comment.parentId)
				this.#replies.set(comment.parentId, [
					...(this.#replies.get(comment.parentId) ?? []),
					comment,
				]);
		});
	}

	get size(): number {
		return this.#comments.size;
	}

	/** A comment that starts a thread (replies are shown in their parent's balloon). */
	thread(id: string): Comment | undefined {
		const comment = this.#comments.get(id)?.comment;
		return comment && !comment.parentId ? comment : undefined;
	}

	/** The first comment of `ids` that is known. */
	first(ids: readonly string[] | undefined): Comment | undefined {
		for (const id of ids ?? []) {
			const comment = this.#comments.get(id)?.comment;
			if (comment) return comment;
		}
		return undefined;
	}

	/** A comment and its replies: label and text, then author and date in grey. */
	paragraphs(comment: Comment, parent?: number): BalloonPiece[][] {
		const number = this.#comments.get(comment.id)?.number ?? 0;
		const color = comment.resolved ? { color: RESOLVED } : {};
		const label = `Commented [${initialsOf(comment)}${number}${parent ? `R${parent}` : ''}]: `;
		const meta = [comment.author, shortDate(comment.date), comment.resolved ? 'Resolved' : '']
			.filter(Boolean)
			.join(', ');
		const own: BalloonPiece[][] = [
			[
				{ text: label, bold: true, family: this.fonts.label, ...color },
				{ text: comment.text, family: this.fonts.text, ...color },
			],
			...(meta ? [[{ text: meta, color: META, family: this.fonts.label }]] : []),
		];
		const replies = (this.#replies.get(comment.id) ?? []).flatMap((reply) =>
			this.paragraphs(reply, number),
		);
		return [...own, ...replies];
	}
}
