/**
 * Internal links, the outline and the structure tags: a link into the document becomes a link to
 * a bookmark placed at the paragraph it lands on; outline entries and tagged headings (H1..H6)
 * set the heading level of the paragraphs they name.
 */
import type { PdfOutlineItem, PdfPageLink } from '../../pdf/index';
import type { Line } from './lines';
import type { HeadingLevel, ParagraphBlock } from './paragraphs';

/** Marks the fragments under an internal link with its destination. */
export function applyInternalLinks(lines: Line[], links: readonly PdfPageLink[]): void {
	for (const link of links) {
		if (link.uri || link.page === undefined) continue;
		const target = { page: link.page, ...(link.targetY !== undefined ? { y: link.targetY } : {}) };
		for (const line of lines) {
			if (line.baseline < link.y || line.baseline > link.y + link.height + line.size * 0.3)
				continue;
			for (const fragment of line.fragments) {
				const middle = fragment.x + fragment.width / 2;
				if (middle >= link.x && middle <= link.x + link.width) fragment.target = target;
			}
		}
	}
}

/** The paragraph holding the first line at or below `y` on `page` (its first line without `y`). */
function paragraphAt(
	blocks: ParagraphBlock[],
	page: number,
	y: number | undefined,
): ParagraphBlock | undefined {
	let best: { block: ParagraphBlock; line: Line } | undefined;
	for (const block of blocks)
		for (const line of block.lines) {
			if (line.page !== page || (y !== undefined && line.bottom < y - 1)) continue;
			if (!best || line.page < best.line.page || line.baseline < best.line.baseline)
				best = { block, line };
		}
	return best?.block;
}

const normalise = (text: string) =>
	text
		.normalize('NFC')
		.replace(/\s+/g, ' ')
		.replace(/[\s.\u2026\d]+$/, '')
		.trim()
		.toLowerCase();
const textOf = (block: ParagraphBlock) => block.fragments.map((fragment) => fragment.text).join('');

/** Gives each internal link a bookmark at the paragraph it goes to. */
export function resolveAnchors(blocks: ParagraphBlock[]): number {
	const names = new Map<ParagraphBlock, string>();
	let unresolved = 0;
	for (const block of blocks)
		for (const fragment of block.fragments) {
			const target = fragment.target;
			if (!target) continue;
			delete fragment.target;
			if (fragment.note) continue;
			const destination = paragraphAt(blocks, target.page, target.y);
			if (!destination) {
				unresolved++;
				continue;
			}
			let name = names.get(destination);
			if (!name) {
				name = `_PdfTarget${names.size + 1}`;
				names.set(destination, name);
				(destination.bookmarks ??= []).push(name);
			}
			fragment.anchor = name;
		}
	return unresolved;
}

/** Heading levels from tagged headings (H1..H6) and from the outline entries that name a paragraph. */
export function applyOutline(blocks: ParagraphBlock[], outline: readonly PdfOutlineItem[]): void {
	for (const block of blocks) {
		const roles = new Set(
			block.fragments.filter((fragment) => fragment.text.trim()).map((fragment) => fragment.role),
		);
		const [role] = roles;
		const tagged = roles.size === 1 && role ? /^H([1-6])$/.exec(role) : null;
		if (tagged) block.heading = Number(tagged[1]) as HeadingLevel;
		else if (roles.size === 1 && role === 'P' && block.heading && block.lines.length > 1)
			delete block.heading;
	}
	for (const item of outline) {
		if (item.page === undefined || !item.title) continue;
		const block = paragraphAt(blocks, item.page, item.y);
		if (!block) continue;
		const title = normalise(item.title);
		const text = normalise(textOf(block));
		if (!title || !text || !(text === title || text.startsWith(title) || title.startsWith(text)))
			continue;
		const fromTags = block.fragments.some(
			(fragment) => fragment.role && /^H[1-6]$/.test(fragment.role),
		);
		if (!fromTags) block.heading = Math.min(6, item.level + 1) as HeadingLevel;
	}
}
