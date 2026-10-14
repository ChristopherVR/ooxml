import type { Node as DocumentNode } from 'prosemirror-model';
export interface PrintParagraphPosition {
	id: string;
	start: number;
	end: number;
}
/** Index immutable editor documents and immutable page render roots independently. */
export class PrintEditingIndex {
	private doc: DocumentNode | undefined;
	private root: Element | null | undefined;
	private renderVersion: string | undefined;
	private paragraphs: PrintParagraphPosition[] = [];
	private positions = new Map<string, PrintParagraphPosition>();
	private byBlock = new Map<string, HTMLElement[]>();
	private lines: HTMLElement[] = [];
	private renderedLines: HTMLElement[] = [];
	private pageLines = new Map<HTMLElement, HTMLElement[]>();
	private linePositions = new Map<HTMLElement, number>();
	constructor(private readonly pages: HTMLElement) {}
	update(doc: DocumentNode): void {
		const documentChanged = this.doc !== doc;
		if (documentChanged) {
			this.doc = doc;
			this.paragraphs = [];
			this.positions.clear();
			doc.descendants((node, pos) => {
				if (node.type.name !== 'paragraph') return true;
				const paragraph = {
					id: String(node.attrs.id),
					start: pos + 1,
					end: pos + 1 + node.content.size,
				};
				this.paragraphs.push(paragraph);
				if (!this.positions.has(paragraph.id)) this.positions.set(paragraph.id, paragraph);
				return false;
			});
		}
		const root = this.pages.firstElementChild;
		const renderVersion = (root as HTMLElement | null)?.dataset.renderVersion;
		const renderChanged = root !== this.root || renderVersion !== this.renderVersion;
		if (!documentChanged && !renderChanged) return;
		if (renderChanged)
			this.renderedLines = [
				...this.pages.querySelectorAll<HTMLElement>('.dve-print-line[data-block-id]'),
			];
		this.root = root;
		this.renderVersion = renderVersion;
		this.byBlock.clear();
		this.pageLines.clear();
		this.linePositions.clear();
		this.lines = [];
		for (const line of this.renderedLines) {
			const id = line.dataset.blockId!;
			if (!this.positions.has(id)) continue;
			const block = this.byBlock.get(id) ?? [];
			block.push(line);
			this.byBlock.set(id, block);
			const page = line.closest<HTMLElement>('.dve-print-page') ?? this.pages;
			const pageRows = this.pageLines.get(page) ?? [];
			pageRows.push(line);
			this.pageLines.set(page, pageRows);
			this.linePositions.set(line, this.lines.length);
			this.lines.push(line);
		}
	}
	position(id: string, offset: number): number | undefined {
		const paragraph = this.positions.get(id);
		return paragraph
			? paragraph.start + Math.max(0, Math.min(offset, paragraph.end - paragraph.start))
			: undefined;
	}
	private firstEndingAtOrAfter(pos: number): number {
		let low = 0,
			high = this.paragraphs.length;
		while (low < high) {
			const mid = (low + high) >>> 1;
			if (this.paragraphs[mid]!.end < pos) low = mid + 1;
			else high = mid;
		}
		return low;
	}
	*between(from: number, to: number): Iterable<PrintParagraphPosition> {
		for (let i = this.firstEndingAtOrAfter(from); i < this.paragraphs.length; i++) {
			const paragraph = this.paragraphs[i]!;
			if (paragraph.start > to) break;
			yield paragraph;
		}
	}
	at(pos: number): PrintParagraphPosition | undefined {
		const paragraph = this.paragraphs[this.firstEndingAtOrAfter(pos)];
		return paragraph && paragraph.start <= pos ? paragraph : undefined;
	}
	block(id: string): readonly HTMLElement[] {
		return this.byBlock.get(id) ?? [];
	}
	adjacent(line: HTMLElement, direction: number): HTMLElement {
		return this.lines[(this.linePositions.get(line) ?? -1) + direction] ?? line;
	}
	pageStep(line: HTMLElement, direction: number): HTMLElement {
		const page = line.closest<HTMLElement>('.dve-print-page') ?? this.pages;
		const pages = [...this.pageLines.keys()];
		return this.pageLines.get(pages[pages.indexOf(page) + direction]!)?.[0] ?? line;
	}
	nearest(x: number, y: number): HTMLElement | undefined {
		const distance = (rect: DOMRect) =>
			Math.hypot(
				Math.max(rect.left - x, x - rect.right, 0),
				Math.max(rect.top - y, y - rect.bottom, 0),
			);
		let bestPage: HTMLElement | undefined,
			bestDistance = Infinity;
		for (const page of this.pageLines.keys()) {
			const score = distance(page.getBoundingClientRect());
			if (score < bestDistance) {
				bestPage = page;
				bestDistance = score;
			}
		}
		let best: HTMLElement | undefined;
		bestDistance = Infinity;
		for (const line of this.pageLines.get(bestPage!) ?? []) {
			const score = distance(line.getBoundingClientRect());
			if (score < bestDistance) {
				best = line;
				bestDistance = score;
			}
		}
		return best;
	}

	clear(): void {
		this.doc = undefined;
		this.root = undefined;
		this.paragraphs = [];
		this.positions.clear();
		this.pageLines.clear();
		this.byBlock.clear();
		this.lines = [];
		this.renderedLines = [];
		this.linePositions.clear();
	}
}
