// Format Cells > Border: line style and colour, the None / Outline / Inside presets, edge toggles
// and a preview. Cell mode applies through the session's border presets; dxf mode writes the
// four outer edges of the differential style.
import type { Border, BorderEdge, BorderStyle, CellRange } from '@christophervr/xlsx-core';
import { resolveColor } from '@christophervr/xlsx-core';
import { type Target, UNSET } from '../../commands/util.js';
import { el, fieldset, row } from '../fields.js';
import { swatchGrid } from './color-swatches.js';
import type { FormatTab, TabInit } from './types.js';

type EdgeId =
	| 'top'
	| 'bottom'
	| 'left'
	| 'right'
	| 'insideH'
	| 'insideV'
	| 'diagonalUp'
	| 'diagonalDown';
const OUTER = ['top', 'bottom', 'left', 'right'] as const;

export const LINE_STYLES: ReadonlyArray<readonly [BorderStyle | 'none', string, string]> = [
	['none', 'None', 'none'],
	['hair', 'Hair', '1px dotted'],
	['dotted', 'Dotted', '1px dotted'],
	['dashDotDot', 'Dash Dot Dot', '1px dotted'],
	['dashDot', 'Dash Dot', '1px dashed'],
	['dashed', 'Dashed', '1px dashed'],
	['thin', 'Thin', '1px solid'],
	['mediumDashDotDot', 'Medium Dash Dot Dot', '2px dotted'],
	['slantDashDot', 'Slanted Dash Dot', '2px dashed'],
	['mediumDashDot', 'Medium Dash Dot', '2px dashed'],
	['mediumDashed', 'Medium Dashed', '2px dashed'],
	['medium', 'Medium', '2px solid'],
	['thick', 'Thick', '3px solid'],
	['double', 'Double', '3px double'],
];

const EDGES: ReadonlyArray<readonly [EdgeId, string]> = [
	['top', 'Top Border'],
	['insideH', 'Inside Horizontal Border'],
	['bottom', 'Bottom Border'],
	['diagonalUp', 'Diagonal Up Border'],
	['left', 'Left Border'],
	['insideV', 'Inside Vertical Border'],
	['right', 'Right Border'],
	['diagonalDown', 'Diagonal Down Border'],
];

const cssOf = (style: BorderStyle): string =>
	LINE_STYLES.find(([s]) => s === style)?.[2] ?? '1px solid';

export function borderTab(init: TabInit): FormatTab {
	const { ctx } = init;
	const b: Border = init.style.border;
	const theme = ctx.workbook()?.theme ?? { colors: [], majorFont: '', minorFont: '' };
	const edges: Partial<Record<EdgeId, BorderEdge | null>> = {
		top: b.top ?? null,
		bottom: b.bottom ?? null,
		left: b.left ?? null,
		right: b.right ?? null,
	};
	if (b.diagonalUp && b.diagonal) edges.diagonalUp = b.diagonal;
	if (b.diagonalDown && b.diagonal) edges.diagonalDown = b.diagonal;
	const changed = new Set<EdgeId>();
	let cleared = false;
	let lineStyle: BorderStyle | 'none' = 'thin';
	const color = swatchGrid(ctx, 'Color:', 'Automatic', undefined);
	const current = (): BorderEdge | null => {
		if (lineStyle === 'none') return null;
		const c = color.value();
		return c ? { style: lineStyle, color: c } : { style: lineStyle };
	};
	const styleButtons = el(ctx, 'div', 'xve-line-styles');
	styleButtons.setAttribute('role', 'radiogroup');
	styleButtons.setAttribute('aria-label', ctx.t('Style:'));
	const styleList: HTMLButtonElement[] = [];
	for (const [value, label, css] of LINE_STYLES) {
		const button = el(ctx, 'button', 'xve-btn xve-line-style');
		button.type = 'button';
		button.setAttribute('role', 'radio');
		button.setAttribute('aria-label', ctx.t(label));
		button.dataset.style = value;
		const line = el(ctx, 'span');
		line.style.display = 'inline-block';
		line.style.width = '48px';
		if (value === 'none') line.textContent = ctx.t(label);
		else line.style.borderTop = `${css} currentColor`;
		button.append(line);
		button.addEventListener('click', () => {
			lineStyle = value;
			paintStyles();
		});
		styleList.push(button);
		styleButtons.append(button);
	}
	const paintStyles = (): void => {
		for (const button of styleList)
			button.setAttribute('aria-checked', String(button.dataset.style === lineStyle));
	};
	const preview = el(ctx, 'div', 'xve-border-preview');
	preview.dataset.preview = '';
	const insideH = el(ctx, 'div');
	const insideV = el(ctx, 'div');
	const svgNs = 'http://www.w3.org/2000/svg';
	const diag = ctx.host.ownerDocument.createElementNS(svgNs, 'svg');
	diag.setAttribute('viewBox', '0 0 120 80');
	diag.setAttribute('style', 'position:absolute;inset:0;width:100%;height:100%');
	Object.assign(insideH.style, { position: 'absolute', left: '0', right: '0', top: '50%' });
	Object.assign(insideV.style, { position: 'absolute', top: '0', bottom: '0', left: '50%' });
	preview.append(insideH, insideV, diag);
	const edgeButtons = new Map<EdgeId, HTMLButtonElement>();
	const css = (edge: BorderEdge | null | undefined): string =>
		edge
			? `${cssOf(edge.style)} ${resolveColor(edge.color, theme, '#000000') ?? '#000000'}`
			: 'none';
	const paint = (): void => {
		for (const side of OUTER) preview.style.setProperty(`border-${side}`, css(edges[side]));
		insideH.style.borderTop = css(edges.insideH);
		insideV.style.borderLeft = css(edges.insideV);
		diag.replaceChildren();
		for (const [id, y1, y2] of [
			['diagonalUp', 80, 0],
			['diagonalDown', 0, 80],
		] as const) {
			const edge = edges[id];
			if (!edge) continue;
			const line = ctx.host.ownerDocument.createElementNS(svgNs, 'line');
			for (const [k, v] of Object.entries({
				x1: 0,
				y1,
				x2: 120,
				y2,
				stroke: resolveColor(edge.color, theme, '#000000') ?? '#000000',
			}))
				line.setAttribute(k, String(v));
			diag.append(line);
		}
		for (const [id, button] of edgeButtons)
			button.setAttribute('aria-pressed', String(!!edges[id]));
	};
	const set = (ids: readonly EdgeId[], edge: BorderEdge | null): void => {
		for (const id of ids) {
			edges[id] = edge;
			changed.add(id);
		}
		paint();
	};
	const inner = !init.dxf;
	const edgeRow = el(ctx, 'div', 'xve-row');
	for (const [id, label] of EDGES) {
		if (!inner && !(OUTER as readonly string[]).includes(id)) continue;
		const button = el(ctx, 'button', 'xve-btn');
		button.type = 'button';
		button.textContent = ctx.t(label);
		button.addEventListener('click', () => set([id], edges[id] ? null : current()));
		edgeButtons.set(id, button);
		edgeRow.append(button);
	}
	const presetButton = (label: string, run: () => void): HTMLButtonElement => {
		const button = el(ctx, 'button', 'xve-btn');
		button.dataset.preset = label.toLowerCase();
		button.type = 'button';
		button.textContent = ctx.t(label);
		button.addEventListener('click', run);
		return button;
	};
	const presets = el(ctx, 'div', 'xve-row');
	presets.append(
		presetButton('None', () => {
			for (const id of edgeButtons.keys()) edges[id] = null;
			changed.clear();
			cleared = true;
			if (init.dxf) for (const id of OUTER) changed.add(id);
			paint();
		}),
		presetButton('Outline', () => set(OUTER, current())),
	);
	if (inner) presets.append(presetButton('Inside', () => set(['insideH', 'insideV'], current())));
	paintStyles();
	paint();
	const panel = el(ctx, 'div');
	panel.append(
		row(ctx, fieldset(ctx, 'Line', styleButtons), fieldset(ctx, 'Color:', color.element)),
		fieldset(ctx, 'Presets', presets),
		fieldset(ctx, 'Border', edgeRow, preview),
	);
	return {
		id: 'border',
		label: 'Border',
		panel,
		dirty: () => cleared || changed.size > 0,
		apply: (t) => {
			for (const r of t.ranges) applyBorders(t, r, edges, changed, cleared);
		},
		toDxf: (dxf) => {
			const next: Border = { ...dxf.border };
			for (const side of OUTER) {
				if (!changed.has(side)) continue;
				const edge = edges[side];
				if (edge) next[side] = edge;
				else delete next[side];
			}
			if (Object.keys(next).length) dxf.border = next;
			else delete dxf.border;
		},
		focus: () => styleList[0]?.focus(),
	};
}

const band = (r: CellRange, rows: [number, number], cols: [number, number]): CellRange => ({
	start: { row: rows[0], col: cols[0] },
	end: { row: rows[1], col: cols[1] },
});

function applyBorders(
	t: Target,
	r: CellRange,
	edges: Partial<Record<EdgeId, BorderEdge | null>>,
	changed: Set<EdgeId>,
	cleared: boolean,
): void {
	const { session, sheet } = t;
	const clear = (range: CellRange, border: Partial<Border>) =>
		session.applyStyle(sheet, [range], { border });
	const { start, end } = r;
	if (cleared) session.setBorders(sheet, r, 'none');
	for (const id of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const) {
		if (!changed.has(id)) continue;
		const edge = edges[id];
		if (edge) {
			session.setBorders(sheet, r, id, edge);
			continue;
		}
		if (id === 'top') clear(band(r, [start.row, start.row], [start.col, end.col]), { top: UNSET });
		if (id === 'bottom')
			clear(band(r, [end.row, end.row], [start.col, end.col]), { bottom: UNSET });
		if (id === 'left')
			clear(band(r, [start.row, end.row], [start.col, start.col]), { left: UNSET });
		if (id === 'right') clear(band(r, [start.row, end.row], [end.col, end.col]), { right: UNSET });
		if (id === 'insideH' && end.row > start.row) {
			clear(band(r, [start.row + 1, end.row], [start.col, end.col]), { top: UNSET });
			clear(band(r, [start.row, end.row - 1], [start.col, end.col]), { bottom: UNSET });
		}
		if (id === 'insideV' && end.col > start.col) {
			clear(band(r, [start.row, end.row], [start.col + 1, end.col]), { left: UNSET });
			clear(band(r, [start.row, end.row], [start.col, end.col - 1]), { right: UNSET });
		}
	}
	if (changed.has('diagonalUp') || changed.has('diagonalDown')) {
		const up = edges.diagonalUp;
		const down = edges.diagonalDown;
		const diagonal = up ?? down;
		session.applyStyle(sheet, [r], {
			border: {
				diagonal: diagonal ?? UNSET,
				diagonalUp: up ? true : UNSET,
				diagonalDown: down ? true : UNSET,
			},
		});
	}
}
