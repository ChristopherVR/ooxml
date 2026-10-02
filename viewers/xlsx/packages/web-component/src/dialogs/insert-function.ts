// Insert Function: search, category filter, the function list and its syntax and description
// from the core FUNCTION_CATALOG, and optional arguments.
import { FUNCTION_CATALOG, type FunctionInfo } from '@christophervr/xlsx-core';
import { insertFunctionCall, recentFunctions } from '../commands/formulas.js';
import type { EditorContext } from '../context.js';
import { el, field, listBox, row, select, text, textInput } from './fields.js';
import { button, showDialog } from './frame.js';

export interface InsertFunctionProps {
	category?: string;
	fn?: string;
}

export const FUNCTION_CATEGORIES = (): string[] =>
	[...new Set(FUNCTION_CATALOG.map((f) => f.category))].sort();

/** Functions matching a search (name first, then description words), like Excel's Go. */
export function searchFunctions(query: string): FunctionInfo[] {
	const q = query.trim().toLowerCase();
	if (!q) return [...FUNCTION_CATALOG];
	const words = q.split(/\s+/);
	const byName = FUNCTION_CATALOG.filter((f) => f.name.toLowerCase().includes(q));
	const byText = FUNCTION_CATALOG.filter(
		(f) => !byName.includes(f) && words.every((w) => f.description.toLowerCase().includes(w)),
	);
	return [...byName, ...byText];
}

export function openInsertFunction(
	ctx: EditorContext,
	props: InsertFunctionProps = {},
): Promise<string | undefined> {
	const search = textInput(ctx);
	search.placeholder = ctx.t('Type a brief description of what you want to do and then click Go');
	const go = button(ctx, 'Go');
	const initial =
		props.category === 'all'
			? 'all'
			: props.category && props.category !== 'recent'
				? props.category
				: 'recent';
	const category = select(
		ctx,
		[
			['recent', 'Most Recently Used'],
			['all', 'All'],
			...FUNCTION_CATEGORIES().map((c) => [c, c] as const),
		],
		initial,
	);
	const syntax = el(ctx, 'p', 'xve-syntax');
	syntax.style.fontWeight = '600';
	const description = el(ctx, 'p', 'xve-note');
	const args = textInput(ctx);
	let confirm: (() => void) | undefined;
	const describe = (name: string): void => {
		const info = FUNCTION_CATALOG.find((f) => f.name === name);
		syntax.textContent = info?.syntax ?? '';
		description.textContent = info?.description ?? '';
	};
	const list = listBox(ctx, 'Select a function:', describe, () => confirm?.());
	const fill = (items: readonly FunctionInfo[]): void => {
		list.setItems(items.map((f) => [f.name, f.name] as const));
		const first = items[0];
		if (first) list.select(first.name);
		else describe('');
	};
	const byCategory = (): FunctionInfo[] => {
		const c = category.value;
		if (c === 'all') return [...FUNCTION_CATALOG];
		if (c === 'recent')
			return recentFunctions().flatMap((name) => FUNCTION_CATALOG.filter((f) => f.name === name));
		return FUNCTION_CATALOG.filter((f) => f.category === c);
	};
	category.addEventListener('change', () => fill(byCategory()));
	const runSearch = (): void => fill(searchFunctions(search.value));
	go.addEventListener('click', runSearch);
	search.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			event.stopPropagation();
			runSearch();
		}
	});
	fill(byCategory());
	if (props.fn && FUNCTION_CATALOG.some((f) => f.name === props.fn)) {
		if (!byCategory().some((f) => f.name === props.fn)) {
			category.value = 'all';
			fill(byCategory());
		}
		list.select(props.fn);
	}
	return showDialog<string>(
		ctx,
		{
			name: 'insert-function',
			heading: 'Insert Function',
			wide: true,
			body: [
				row(ctx, field(ctx, 'Search for a function:', search), go),
				field(ctx, 'Or select a category:', category),
				text(ctx, 'Select a function:', 'xve-field-label'),
				list.element,
				syntax,
				description,
				field(ctx, 'Arguments (optional):', args),
			],
			opened: () => search.focus(),
			submit: () => {
				const name = list.value();
				if (!name) return undefined;
				const typed = args.value.trim();
				insertFunctionCall(ctx, name, typed ? typed : undefined);
				return name;
			},
		},
		(open) => {
			confirm = () => open.element.querySelector<HTMLButtonElement>('.xve-btn-primary')?.click();
		},
	);
}
