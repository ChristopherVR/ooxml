/**
 * A thin authoring layer over the core's Word document model: every block, run, note and comment a
 * sample holds is an `ooxml-core` model object that `saveDocx` writes. The only thing added here is
 * the sentinel mechanism (`rawRun`, `rawBlock`) for markup the model cannot express yet, which
 * `finish.ts` swaps in after the core has saved.
 */
import type {
	Comment,
	InlineImage,
	Note,
	Paragraph,
	PendingMediaPart,
	PicturePlacement,
	Revision,
	TextRun,
} from '../../../../src/core/docx/index';

export type Runs = TextRun[];
/** Anything a paragraph can be made of: plain text, one run, or runs. */
export type Inline = string | TextRun | Runs;

export const toRuns = (content: Inline): Runs =>
	typeof content === 'string' ? [{ text: content }] : Array.isArray(content) ? content : [content];

/** One run of text with direct formatting. */
export const t = (text: string, props: Partial<TextRun> = {}): TextRun => ({ ...props, text });

/**
 * Runs from a line of light markup: `**bold**`, `*italic*`, `^superscript^`, `~subscript~` and
 * `{name}` for the runs `inserts` supplies (fields, notes, equations). `base` is applied to every
 * text run the markup yields, not to the inserts.
 */
export function rich(
	markup: string,
	inserts: Record<string, Inline> = {},
	base: Partial<TextRun> = {},
): Runs {
	const runs: Runs = [];
	const state = { bold: false, italic: false, sup: false, sub: false };
	let text = '';
	const flush = () => {
		if (!text) return;
		runs.push({
			...base,
			text,
			...(state.bold && { bold: true }),
			...(state.italic && { italic: true }),
			...(state.sup && { verticalAlign: 'superscript' as const }),
			...(state.sub && { verticalAlign: 'subscript' as const }),
		});
		text = '';
	};
	for (let at = 0; at < markup.length; at++) {
		const char = markup[at]!;
		if (char === '*') {
			flush();
			if (markup[at + 1] === '*') {
				state.bold = !state.bold;
				at++;
			} else state.italic = !state.italic;
		} else if (char === '^') {
			flush();
			state.sup = !state.sup;
		} else if (char === '~') {
			flush();
			state.sub = !state.sub;
		} else if (char === '{') {
			const end = markup.indexOf('}', at);
			const name = markup.slice(at + 1, end);
			const insert = inserts[name];
			if (insert === undefined) throw new Error(`rich(): no insert named "${name}"`);
			flush();
			runs.push(...toRuns(insert));
			at = end;
		} else text += char;
	}
	flush();
	return runs;
}

export interface PictureOptions {
	/** File name stem under `word/media/`. */
	name: string;
	widthPx: number;
	heightPx: number;
	alt: string;
	title?: string;
	/** Floating placement; absent for an inline picture. */
	placement?: PicturePlacement;
}

export interface Reviewer {
	author: string;
	initials: string;
}

/** The parts of one document as they are collected, before `saveDocx` and `finish`. */
export class Doc {
	readonly footnotes: Note[] = [];
	readonly endnotes: Note[] = [];
	readonly comments: Comment[] = [];
	readonly media = new Map<string, PendingMediaPart>();
	/** Sentinel text to the XML that replaces its run (or its paragraph) after the core saved. */
	readonly raw = new Map<string, string>();
	private count = 0;
	private markId = 20000;
	private revisionCount = 0;

	/** A block id unique in the document, as the model requires. */
	id(prefix = 'p'): string {
		return `${prefix}${++this.count}`;
	}

	/** An id for a bookmark written by hand, clear of the ones the core allocates. */
	nextMarkId(): number {
		return ++this.markId;
	}

	/** A paragraph of `style` holding `content`. */
	p(style: string, content: Inline = '', props: Partial<Paragraph> = {}): Paragraph {
		return { type: 'paragraph', id: this.id(), style, runs: toRuns(content), ...props };
	}

	/** A run the finishing pass replaces with `xml` (inline markup the model has no field for). */
	rawRun(xml: string): TextRun {
		const key = `[[RAW:${this.raw.size + 1}]]`;
		this.raw.set(key, xml);
		return { text: key };
	}

	/** A paragraph the finishing pass replaces with `xml` (block markup the model cannot write). */
	rawBlock(xml: string): Paragraph {
		const key = `[[RAWBLOCK:${this.raw.size + 1}]]`;
		this.raw.set(key, xml);
		return { type: 'paragraph', id: this.id(), runs: [{ text: key }] };
	}

	/**
	 * `runs` inside a bookmark that covers exactly them. The model keeps bookmarks per paragraph
	 * (`Paragraph.bookmarks` brackets the whole paragraph), so a partial range is written by hand.
	 */
	bookmark(name: string, runs: Inline): Runs {
		const id = this.nextMarkId();
		return [
			this.rawRun(`<w:bookmarkStart w:id="${id}" w:name="${name}"/>`),
			...toRuns(runs),
			this.rawRun(`<w:bookmarkEnd w:id="${id}"/>`),
		];
	}

	private note(kind: 'footnote' | 'endnote', content: Inline): TextRun {
		const notes = kind === 'footnote' ? this.footnotes : this.endnotes;
		const id = String(notes.length + 1);
		const name = kind === 'footnote' ? 'Footnote' : 'Endnote';
		notes.push({
			id,
			blocks: [
				{
					type: 'paragraph',
					id: this.id(kind === 'footnote' ? 'fn' : 'en'),
					style: `${name}Text`,
					runs: [
						{ text: '', noteMark: kind, style: `${name}Reference` },
						{ text: ' ' },
						...toRuns(content),
					],
				},
			],
		});
		return { text: '', noteReference: { kind, id }, style: `${name}Reference` };
	}

	/** Adds a footnote and returns its reference mark. */
	footnote(content: Inline): TextRun {
		return this.note('footnote', content);
	}

	/** Adds an endnote and returns its reference mark. */
	endnote(content: Inline): TextRun {
		return this.note('endnote', content);
	}

	/**
	 * Anchors a comment thread on `runs`: the first entry is the comment, the rest are replies.
	 * Dates are fixed text, so nothing depends on the clock.
	 */
	comment(runs: Inline, thread: readonly (Reviewer & { date: string; text: string })[]): Runs {
		const [first, ...replies] = thread;
		if (!first) return toRuns(runs);
		const id = `c${this.comments.length + 1}`;
		this.comments.push({
			id,
			author: first.author,
			initials: first.initials,
			date: first.date,
			text: first.text,
		});
		for (const reply of replies)
			this.comments.push({
				id: `c${this.comments.length + 1}`,
				author: reply.author,
				initials: reply.initials,
				date: reply.date,
				text: reply.text,
				parentId: id,
			});
		return toRuns(runs).map((run) => ({ ...run, commentIds: [...(run.commentIds ?? []), id] }));
	}

	private revision(kind: Revision['kind'], reviewer: Reviewer, date: string): Revision {
		// The core renumbers revision ids on save; these only need to be distinct.
		return { kind, author: reviewer.author, date, id: `rev${++this.revisionCount}` };
	}

	/** `runs` as a tracked insertion. */
	inserted(runs: Inline, reviewer: Reviewer, date: string): Runs {
		return toRuns(runs).map((run) => ({
			...run,
			revision: this.revision('insert', reviewer, date),
		}));
	}

	/** `runs` as a tracked deletion. */
	deleted(runs: Inline, reviewer: Reviewer, date: string): Runs {
		return toRuns(runs).map((run) => ({
			...run,
			revision: this.revision('delete', reviewer, date),
		}));
	}

	/** `runs` with a tracked formatting change; `previousXml` is the `w:rPr` they had before. */
	reformatted(runs: Inline, reviewer: Reviewer, date: string, previousXml = W_RPR_EMPTY): Runs {
		return toRuns(runs).map((run) => ({
			...run,
			revision: {
				...this.revision('formatChange', reviewer, date),
				previousRunPropertiesXml: previousXml,
			},
		}));
	}

	/** A picture run: the bytes become a media part when the core saves. */
	picture(bytes: Uint8Array, options: PictureOptions): TextRun {
		const partName = `word/media/${options.name}.png`;
		this.media.set(partName, { bytes, contentType: 'image/png' });
		const image: InlineImage = {
			relId: '',
			partName,
			contentType: 'image/png',
			widthPx: options.widthPx,
			heightPx: options.heightPx,
			altText: options.alt,
			drawingName: options.title ?? options.name,
			...(options.title && { title: options.title }),
			...(options.placement && { anchored: true, placement: options.placement }),
		};
		return { text: '', image };
	}
}

export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W_RPR_EMPTY = `<w:rPr xmlns:w="${W_NS}"/>`;
