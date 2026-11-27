// Tracked-change author colours, shared by the editor's review marks and the PDF export. DOM-free.

/** Stable, accessible palette for attributing tracked-change authors by name. */
const REVIEW_AUTHOR_PALETTE = [
	'#c2431f',
	'#1a73e8',
	'#188038',
	'#8430ce',
	'#b2103c',
	'#00838f',
	'#e37400',
	'#5b6bc0',
];
export function authorColor(author: string): string {
	let hash = 0;
	for (let i = 0; i < author.length; i++) hash = (hash * 31 + author.charCodeAt(i)) >>> 0;
	return REVIEW_AUTHOR_PALETTE[hash % REVIEW_AUTHOR_PALETTE.length] ?? '#c2431f';
}
