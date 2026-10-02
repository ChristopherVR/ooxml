// Format Cells > Font: name, style, size, underline, colour, effects and a live preview. In dxf
// (conditional format) mode the name and size are not offered, as in Excel.
import type { Color, Font, UnderlineStyle } from '@christophervr/xlsx-core';
import { resolveColor } from '@christophervr/xlsx-core';
import { FONT_NAMES, FONT_SIZES } from '../../commands/font.js';
import { UNSET } from '../../commands/util.js';
import { checkbox, el, field, fieldset, nextId, row, select, textInput } from '../fields.js';
import { swatchGrid } from './color-swatches.js';
import type { FormatTab, TabInit } from './types.js';

const STYLES: ReadonlyArray<readonly [string, string]> = [
	['regular', 'Regular'],
	['italic', 'Italic'],
	['bold', 'Bold'],
	['boldItalic', 'Bold Italic'],
];

const UNDERLINES: ReadonlyArray<readonly [string, string]> = [
	['none', 'None'],
	['single', 'Single'],
	['double', 'Double'],
	['singleAccounting', 'Single Accounting'],
	['doubleAccounting', 'Double Accounting'],
];

/** An editable text input with a native suggestion list. */
function comboInput(init: TabInit, value: string, values: readonly string[]): HTMLInputElement {
	const input = textInput(init.ctx, value);
	const list = el(init.ctx, 'datalist');
	list.id = nextId('list');
	for (const v of values) {
		const option = el(init.ctx, 'option');
		option.value = v;
		list.append(option);
	}
	input.setAttribute('list', list.id);
	queueMicrotask(() => input.parentElement?.append(list));
	return input;
}

export function fontTab(init: TabInit): FormatTab {
	const { ctx } = init;
	const f: Font = init.style.font;
	const defaults = ctx.workbook()?.styles[0]?.font ?? {};
	const changed = new Set<string>();
	const mark = (key: string) => () => {
		changed.add(key);
		normal.input.checked = false;
		preview();
	};
	const name = comboInput(init, f.name ?? defaults.name ?? 'Calibri', FONT_NAMES);
	const style = select(
		ctx,
		STYLES,
		f.bold ? (f.italic ? 'boldItalic' : 'bold') : f.italic ? 'italic' : 'regular',
	);
	const size = comboInput(init, String(f.size ?? defaults.size ?? 11), FONT_SIZES.map(String));
	const underline = select(ctx, UNDERLINES, f.underline ?? 'none');
	const strike = checkbox(ctx, 'Strikethrough', !!f.strike);
	const superscript = checkbox(ctx, 'Superscript', f.vertAlign === 'superscript');
	const subscript = checkbox(ctx, 'Subscript', f.vertAlign === 'subscript');
	const normal = checkbox(ctx, 'Normal font', false);
	const auto = (c: Color | undefined) => !c || c.auto || (c.theme === 1 && !c.tint);
	const color = swatchGrid(ctx, 'Color:', 'Automatic', auto(f.color) ? undefined : f.color, () =>
		mark('color')(),
	);
	const sample = el(ctx, 'div', 'xve-sample');
	sample.dataset.fontPreview = '';
	sample.textContent = 'AaBbCcYyZz';
	const preview = (): void => {
		const theme = ctx.workbook()?.theme ?? { colors: [], majorFont: '', minorFont: '' };
		const s = style.value;
		sample.style.fontFamily = `"${name.value}"`;
		sample.style.fontSize = `${Math.min(36, Number(size.value) || 11)}pt`;
		sample.style.fontWeight = s.startsWith('bold') ? 'bold' : 'normal';
		sample.style.fontStyle = /italic/i.test(s) ? 'italic' : 'normal';
		sample.style.textDecorationLine =
			[underline.value !== 'none' ? 'underline' : '', strike.input.checked ? 'line-through' : '']
				.join(' ')
				.trim() || 'none';
		sample.style.textDecorationStyle = /double/i.test(underline.value) ? 'double' : 'solid';
		sample.style.color = resolveColor(color.value(), theme) ?? '#000000';
		sample.style.verticalAlign = superscript.input.checked
			? 'super'
			: subscript.input.checked
				? 'sub'
				: '';
	};
	name.addEventListener('input', mark('name'));
	style.addEventListener('change', mark('style'));
	size.addEventListener('input', mark('size'));
	underline.addEventListener('change', mark('underline'));
	strike.input.addEventListener('change', mark('strike'));
	superscript.input.addEventListener('change', () => {
		if (superscript.input.checked) subscript.input.checked = false;
		mark('vertAlign')();
	});
	subscript.input.addEventListener('change', () => {
		if (subscript.input.checked) superscript.input.checked = false;
		mark('vertAlign')();
	});
	normal.input.addEventListener('change', () => {
		if (!normal.input.checked) return;
		name.value = defaults.name ?? 'Calibri';
		size.value = String(defaults.size ?? 11);
		style.value = 'regular';
		underline.value = 'none';
		strike.input.checked = superscript.input.checked = subscript.input.checked = false;
		color.set(undefined);
		for (const key of ['name', 'size', 'style', 'underline', 'strike', 'vertAlign', 'color'])
			changed.add(key);
		preview();
	});
	preview();
	const top = init.dxf
		? row(ctx, field(ctx, 'Font style:', style), field(ctx, 'Underline:', underline))
		: row(
				ctx,
				field(ctx, 'Font:', name),
				field(ctx, 'Font style:', style),
				field(ctx, 'Size:', size),
			);
	const panel = el(ctx, 'div');
	panel.append(
		top,
		...(init.dxf ? [] : [row(ctx, field(ctx, 'Underline:', underline), normal.wrapper)]),
		row(
			ctx,
			fieldset(ctx, 'Color:', color.element),
			fieldset(ctx, 'Effects', strike.wrapper, superscript.wrapper, subscript.wrapper),
		),
		fieldset(ctx, 'Preview', sample),
	);
	/** The changed font keys; `undefined` values clear the key. */
	const changes = (): Partial<Font> => {
		const out: Partial<Font> = {};
		const s = style.value;
		if (changed.has('name') && name.value.trim()) {
			out.name = name.value.trim();
			out.scheme = UNSET;
		}
		if (changed.has('size')) {
			const n = Number(size.value);
			if (Number.isFinite(n) && n >= 1 && n <= 409) out.size = Math.round(n * 2) / 2;
		}
		if (changed.has('style')) {
			out.bold = s.startsWith('bold') || UNSET;
			out.italic = /italic/i.test(s) || UNSET;
		}
		if (changed.has('underline'))
			out.underline = underline.value === 'none' ? UNSET : (underline.value as UnderlineStyle);
		if (changed.has('strike')) out.strike = strike.input.checked || UNSET;
		if (changed.has('vertAlign'))
			out.vertAlign = superscript.input.checked
				? 'superscript'
				: subscript.input.checked
					? 'subscript'
					: UNSET;
		if (changed.has('color')) out.color = color.value() ?? (init.dxf ? UNSET : { theme: 1 });
		return out;
	};
	return {
		id: 'font',
		label: 'Font',
		panel,
		dirty: () => changed.size > 0,
		patch: () => (changed.size ? { font: changes() } : {}),
		toDxf: (dxf) => {
			const next: Font = { ...dxf.font, ...changes() };
			for (const key of Object.keys(next) as (keyof Font)[])
				if (next[key] === undefined) delete next[key];
			if (Object.keys(next).length) dxf.font = next;
			else delete dxf.font;
		},
		focus: () => (init.dxf ? style : name).focus(),
	};
}
