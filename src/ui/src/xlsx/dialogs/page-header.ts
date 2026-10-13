import type { PageSetup } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, panel, row, select, text, textArea } from './fields';
import type { SetupTab } from './page-setup-tabs';

export function headerTab(ctx: EditorContext, setup: PageSetup): SetupTab {
	const changes = new Map<keyof PageSetup, () => string | boolean | undefined>();
	const fields = [
		['header', 'Header:'],
		['footer', 'Footer:'],
		['firstHeader', 'First Page Header:'],
		['firstFooter', 'First Page Footer:'],
		['evenHeader', 'Even Page Header:'],
		['evenFooter', 'Even Page Footer:'],
	] as const;
	let active: HTMLTextAreaElement;
	const inputs = fields.map(([key, label]) => {
		const original = setup[key] ?? '';
		const shown =
			original.startsWith('&C') && !/&[LR]/.test(original.slice(2)) ? original.slice(2) : original;
		const input = textArea(ctx, shown, 2);
		const read = () =>
			input.value ? (/&[LCR]/.test(input.value) ? input.value : `&C${input.value}`) : undefined;
		input.addEventListener('input', () => changes.set(key, read));
		input.addEventListener('change', () => changes.set(key, read));
		input.addEventListener('focus', () => {
			active = input;
		});
		return { key, input, field: field(ctx, label, input) };
	});
	active = inputs[0]!.input;
	const flags = [
		['differentFirst', 'Different first page', false],
		['differentOddEven', 'Different odd and even pages', false],
		['scaleWithDoc', 'Scale header and footer with document', true],
		['alignWithMargins', 'Align header and footer with margins', true],
	] as const;
	const checks = flags.map(([key, label, fallback]) => {
		const check = checkbox(ctx, label, setup[key] ?? fallback);
		check.input.addEventListener('change', () => {
			changes.set(key, () => check.input.checked);
			update();
		});
		return { key, ...check };
	});
	const update = () => {
		for (const input of inputs)
			if (input.key.startsWith('first') || input.key.startsWith('even'))
				input.input.disabled = !checks.find(
					(check) =>
						check.key === (input.key.startsWith('first') ? 'differentFirst' : 'differentOddEven'),
				)!.input.checked;
	};
	update();
	const insert = select(
		ctx,
		[
			['', 'Insert Field'],
			['&P', 'Page Number'],
			['&N', 'Total Pages'],
			['&A', 'Sheet Name'],
			['&F', 'File Name'],
			['&D', 'Date'],
			['&T', 'Time'],
		],
		'',
	);
	const insertCode = (control: HTMLSelectElement, wrap = false) => {
		if (!control.value || active.disabled) return;
		const start = active.selectionStart;
		const selected = wrap ? active.value.slice(start, active.selectionEnd) : '';
		active.setRangeText(
			control.value + selected + (selected ? control.value : ''),
			start,
			active.selectionEnd,
			'end',
		);
		active.dispatchEvent(new ctx.host.ownerDocument.defaultView!.Event('input', { bubbles: true }));
		control.value = '';
		active.focus();
	};
	insert.addEventListener('change', () => insertCode(insert));
	const formatting = select(
		ctx,
		[
			['', 'Insert Formatting'],
			['&B', 'Bold'],
			['&I', 'Italic'],
			['&U', 'Underline'],
			['&S', 'Strikethrough'],
		],
		'',
	);
	formatting.addEventListener('change', () => insertCode(formatting, true));
	return {
		panel: panel(
			ctx,
			row(ctx, field(ctx, 'Insert Field', insert)),
			row(ctx, field(ctx, 'Insert Formatting', formatting)),
			text(
				ctx,
				'Formatting: &B bold, &I italic, &U underline, &S strike, &18 font size, &KFF0000 RGB color, &"Arial,Regular" font.',
			),
			...checks.map((check) => check.wrapper),
			text(
				ctx,
				'Use &L, &C and &R for sections. Fields: &P page, &N total pages, &A sheet, &F file, &D date, &T time. Use && for an ampersand. Unsupported formatting codes are preserved.',
			),
			...inputs.map((input) => input.field),
		),
		read: () =>
			Object.fromEntries([...changes].map(([key, read]) => [key, read()])) as Partial<PageSetup>,
	};
}
