// Symbol: a grid of common characters by subset with a recently used row; Insert adds the
// character to the active cell (through the in-cell editor when the grid is mounted).
import { getCell } from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { el, field, select, text } from './fields.js';
import { showDialog } from './frame.js';

const range = (from: number, to: number, skip: number[] = []): string[] => {
	const out: string[] = [];
	for (let c = from; c <= to; c++) if (!skip.includes(c)) out.push(String.fromCodePoint(c));
	return out;
};

export const SYMBOL_SUBSETS: ReadonlyArray<readonly [id: string, label: string, chars: string[]]> =
	[
		['latin1', 'Latin-1 Supplement', range(0xa1, 0xff, [0xad])],
		[
			'punctuation',
			'General Punctuation',
			[
				'‐',
				'–',
				'‘',
				'’',
				'‚',
				'“',
				'”',
				'„',
				'†',
				'‡',
				'•',
				'…',
				'‰',
				'′',
				'″',
				'‹',
				'›',
				'‼',
				'⁄',
			],
		],
		[
			'currency',
			'Currency Symbols',
			['$', '¢', '£', '¥', '€', '₣', '₤', '₦', '₩', '₪', '₫', '₭', '₱', '₹', '₺', '₽', '₿'],
		],
		[
			'math',
			'Mathematical Operators',
			[
				'±',
				'×',
				'÷',
				'∀',
				'∂',
				'∃',
				'∅',
				'∇',
				'∈',
				'∉',
				'∏',
				'∑',
				'−',
				'√',
				'∞',
				'∠',
				'∧',
				'∨',
				'∩',
				'∪',
				'∫',
				'≈',
				'≠',
				'≡',
				'≤',
				'≥',
				'⊂',
				'⊃',
				'⊆',
				'⊇',
				'⊕',
				'⊥',
				'°',
				'‰',
				'′',
			],
		],
		['arrows', 'Arrows', range(0x2190, 0x21a5).concat(['↵', '⇐', '⇑', '⇒', '⇓', '⇔'])],
		['greek', 'Greek and Coptic', range(0x391, 0x3a9, [0x3a2]).concat(range(0x3b1, 0x3c9))],
		[
			'box',
			'Box Drawing',
			[
				'─',
				'│',
				'┌',
				'┐',
				'└',
				'┘',
				'├',
				'┤',
				'┬',
				'┴',
				'┼',
				'═',
				'║',
				'╔',
				'╗',
				'╚',
				'╝',
				'╠',
				'╣',
				'╦',
				'╩',
				'╬',
				'▀',
				'▄',
				'█',
				'░',
				'▒',
				'▓',
				'■',
				'□',
				'▲',
				'►',
				'▼',
				'◄',
				'○',
				'●',
				'★',
				'☆',
				'✓',
				'✗',
			],
		],
	];

const recent: string[] = [
	'€',
	'£',
	'¥',
	'©',
	'®',
	'™',
	'±',
	'≠',
	'≤',
	'≥',
	'÷',
	'×',
	'∞',
	'µ',
	'α',
	'β',
];

export const codeOf = (ch: string): string =>
	`U+${(ch.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')}`;

/** Inserts `symbol` into the active cell. */
export function insertSymbol(ctx: EditorContext, symbol: string): void {
	const t = target(ctx);
	if (!t) return;
	const cell = getCell(t.ws, t.active.row, t.active.col);
	const current =
		cell?.formula !== undefined
			? `=${cell.formula}`
			: cell?.value === null || cell?.value === undefined
				? ''
				: typeof cell.value === 'object'
					? cell.value.error
					: String(cell.value);
	const grid = ctx.grid();
	if (grid && !grid.isEditing()) grid.beginEdit(current + symbol);
	else t.session.setCellInput(t.sheet, t.active.row, t.active.col, current + symbol);
	const i = recent.indexOf(symbol);
	if (i >= 0) recent.splice(i, 1);
	recent.unshift(symbol);
	recent.length = Math.min(recent.length, 16);
}

export function openSymbol(ctx: EditorContext): Promise<string | undefined> {
	const subset = select(
		ctx,
		SYMBOL_SUBSETS.map(([id, label]) => [id, label] as const),
		'latin1',
	);
	const grid = el(ctx, 'div', 'xve-symbols');
	grid.setAttribute('role', 'grid');
	grid.setAttribute('aria-label', ctx.t('Symbols'));
	const recentRow = el(ctx, 'div', 'xve-symbols');
	recentRow.setAttribute('aria-label', ctx.t('Recently used symbols:'));
	const code = text(ctx, 'Character code: {code}', 'xve-note', { code: '' });
	let chosen = '';
	let confirm: () => void = () => undefined;
	const choose = (ch: string, b?: HTMLButtonElement): void => {
		chosen = ch;
		code.textContent = ctx.t('Character code: {code}', { code: codeOf(ch) });
		for (const other of [...grid.children, ...recentRow.children])
			other.setAttribute('aria-pressed', String(other === b));
	};
	const cellButton = (ch: string): HTMLButtonElement => {
		const b = el(ctx, 'button', 'xve-btn');
		b.type = 'button';
		b.textContent = ch;
		b.setAttribute('aria-label', codeOf(ch));
		b.addEventListener('click', () => choose(ch, b));
		b.addEventListener('dblclick', () => {
			choose(ch, b);
			confirm();
		});
		return b;
	};
	const fill = (): void => {
		const chars = SYMBOL_SUBSETS.find(([id]) => id === subset.value)?.[2] ?? [];
		grid.replaceChildren(...chars.map(cellButton));
		const first = grid.firstElementChild as HTMLButtonElement | null;
		if (first && !chosen) choose(first.textContent ?? '', first);
	};
	subset.addEventListener('change', fill);
	recentRow.replaceChildren(...recent.map(cellButton));
	fill();
	grid.addEventListener('keydown', (event) => {
		const list = [...grid.querySelectorAll<HTMLButtonElement>('button')];
		const i = list.indexOf(event.target as HTMLButtonElement);
		const cols = 16;
		const delta = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: cols, ArrowUp: -cols }[event.key];
		if (event.key === 'Enter' && i >= 0) {
			event.preventDefault();
			choose(list[i]?.textContent ?? '', list[i]);
			confirm();
			return;
		}
		if (i < 0 || delta === undefined) return;
		event.preventDefault();
		const next = list[Math.max(0, Math.min(list.length - 1, i + delta))];
		next?.focus();
		if (next) choose(next.textContent ?? '', next);
	});
	return showDialog<string>(
		ctx,
		{
			name: 'symbol',
			heading: 'Symbol',
			wide: true,
			okLabel: 'Insert',
			body: [
				field(ctx, 'Subset:', subset),
				grid,
				text(ctx, 'Recently used symbols:', 'xve-field-label'),
				recentRow,
				code,
			],
			opened: () => (grid.firstElementChild as HTMLElement | null)?.focus(),
			submit: () => {
				if (!chosen) return undefined;
				insertSymbol(ctx, chosen);
				return chosen;
			},
		},
		(open) => {
			confirm = () => open.element.querySelector<HTMLButtonElement>('.xve-btn-primary')?.click();
		},
	);
}
