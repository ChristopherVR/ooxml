import type { Block, Paragraph, Revision, TextRun } from '../model';
import { DIRECT_RUN_PROPERTY_KEYS } from '../run-formatting';
import { runPropertiesElement } from '../run-properties-snapshot';
import { buildXml } from '../xml';
import type { ChangeKind, CompareContext } from './context';
import type { Piece, Token } from './tokens';

/** One aligned step of a story: a token of the original, of the revised document, or of both. */
export type StreamOp =
	| { kind: 'equal'; a: Token; b: Token }
	| { kind: 'delete'; a: Token }
	| { kind: 'insert'; b: Token };

/** The direct formatting of a run, as compared; text, revisions and identities are not part of it. */
function formatSignature(run: TextRun): string {
	return JSON.stringify(DIRECT_RUN_PROPERTY_KEYS.map((key) => run[key] ?? null));
}

/**
 * Builds the result blocks of one story (the body or a table cell) from its aligned tokens:
 * runs carry `w:ins`/`w:del`/`w:rPrChange` revisions, one revision per contiguous span, and each
 * paragraph mark closes a paragraph with the revised (kept or inserted mark) or original (deleted
 * mark) paragraph's properties.
 */
export class StoryBuilder {
	readonly blocks: Block[] = [];
	private runs: TextRun[] = [];
	private sources: TextRun[] = [];
	private deferred: Block[] = [];
	private span: { kind: ChangeKind; previous?: string; revision: Revision } | undefined;
	/** Original runs as imported, so the deleted pieces of one run merge into one run again. */
	private readonly imported = new Map<TextRun, TextRun | undefined>();

	constructor(private readonly context: CompareContext) {}

	private revisionFor(kind: ChangeKind | undefined, previous?: string): Revision | undefined {
		if (!kind) {
			this.span = undefined;
			return undefined;
		}
		if (this.span?.kind !== kind || this.span.previous !== previous)
			this.span = {
				kind,
				...(previous === undefined ? {} : { previous }),
				revision: this.context.revision(kind, previous),
			};
		return this.span.revision;
	}

	private push(source: TextRun, text: string, kind?: ChangeKind, previous?: string): void {
		const revision = this.revisionFor(kind, previous);
		const last = this.runs.length - 1;
		if (last >= 0 && this.sources[last] === source && this.runs[last]!.revision === revision) {
			this.runs[last]!.text += text;
			return;
		}
		const run: TextRun = { ...source, text };
		delete run.revision;
		delete run.formatRevision;
		if (revision) run.revision = revision;
		this.runs.push(run);
		this.sources.push(source);
	}

	private pieces(token: Token): Piece[] {
		if (token.kind === 'text') return token.pieces;
		if (token.kind === 'atomic') return [{ run: token.run, text: token.run.text }];
		return [];
	}

	/** Unchanged text: the revised runs, with a formatting revision where the formatting differs. */
	private equalText(a: Piece[], b: Piece[]): void {
		let ai = 0;
		let aOffset = 0;
		for (const piece of b) {
			let offset = 0;
			while (offset < piece.text.length) {
				const original = a[ai];
				if (!original) {
					this.push(piece.run, piece.text.slice(offset));
					break;
				}
				const take = Math.min(piece.text.length - offset, original.text.length - aOffset);
				const text = piece.text.slice(offset, offset + take);
				if (
					this.context.compareFormatting &&
					formatSignature(original.run) !== formatSignature(piece.run)
				)
					this.push(piece.run, text, 'formatChange', buildXml(runPropertiesElement(original.run)));
				else this.push(piece.run, text);
				offset += take;
				aOffset += take;
				if (aOffset >= original.text.length) {
					ai++;
					aOffset = 0;
				}
			}
		}
	}

	private block(block: Block): void {
		if (this.runs.length) this.deferred.push(block);
		else this.blocks.push(block);
	}

	private close(paragraph: Paragraph, mark?: 'insert' | 'delete'): void {
		const runs = this.runs;
		const result: Paragraph = { ...paragraph, runs };
		delete result.markRevision;
		delete result.formatRevision;
		if (mark) result.markRevision = this.context.revision(mark);
		this.blocks.push(result, ...this.deferred);
		this.runs = [];
		this.sources = [];
		this.deferred = [];
		this.span = undefined;
	}

	/** Adds one aligned step; tables come in through {@link addBlock}. */
	add(op: StreamOp): void {
		if (op.kind === 'equal') {
			if (op.b.kind === 'mark') return this.close(op.b.paragraph);
			if (op.b.kind === 'text' && op.a.kind === 'text')
				return this.equalText(op.a.pieces, op.b.pieces);
			for (const piece of this.pieces(op.b)) this.push(piece.run, piece.text);
			return;
		}
		if (op.kind === 'insert') {
			if (op.b.kind === 'mark') return this.close(op.b.paragraph, 'insert');
			for (const piece of this.pieces(op.b)) this.push(piece.run, piece.text, 'insert');
			return;
		}
		if (op.a.kind === 'mark') {
			const paragraph = this.context.importParagraph(op.a.paragraph);
			return this.close({ ...paragraph, id: this.context.blockId() }, 'delete');
		}
		for (const piece of this.pieces(op.a)) {
			if (!this.imported.has(piece.run))
				this.imported.set(piece.run, this.context.importRun(piece.run));
			const run = this.imported.get(piece.run);
			if (run) this.push(run, piece.text, 'delete');
		}
	}

	addBlock(block: Block): void {
		this.block(block);
	}

	/** The finished blocks; a story always ends with a paragraph mark, so nothing stays open. */
	finish(): Block[] {
		if (this.runs.length || this.deferred.length)
			throw new Error('A compared story ended inside a paragraph.');
		return this.blocks;
	}
}
