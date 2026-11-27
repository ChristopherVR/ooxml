// Ink through the clipboard: a copied or cut ink drawing carries its strokes, size, position and
// name in data attributes of its placeholder, so pasting it (here or in another editor) brings the
// ink back. The pasted ink has no part of its own yet (no relationship, no part name): the save
// writes it a new one, so a copy never shares the original's part.
import type { Node as ProseMirrorNode } from 'prosemirror-model';

/** Clipboard HTML of other origins may carry anything: past this size it is not read as ink. */
const MAX_CLIPBOARD_INK = 4 * 1024 * 1024;

/** The data attributes an ink drawing's placeholder carries. */
export function inkClipboardAttrs(node: ProseMirrorNode): Record<string, string> {
	if (typeof node.attrs.ink !== 'string') return {};
	let ink: Record<string, unknown>;
	try {
		ink = JSON.parse(node.attrs.ink) as Record<string, unknown>;
	} catch {
		return {};
	}
	const { relId: _relId, partName: _partName, truncated: _truncated, ...own } = ink;
	return {
		'data-ink': JSON.stringify(own),
		'data-unsupported': String(node.attrs.unsupported ?? 'Ink'),
		'data-width-px': String(node.attrs.widthPx),
		'data-height-px': String(node.attrs.heightPx),
		'data-anchored': node.attrs.anchored ? '1' : '0',
		...(typeof node.attrs.placement === 'string' ? { 'data-placement': node.attrs.placement } : {}),
		...(node.attrs.drawingName ? { 'data-drawing-name': String(node.attrs.drawingName) } : {}),
	};
}

/** What a pasted drawing may hold: as much as a drawing this editor collects (`canJoinInk`). */
const MAX_STROKES = 64;
const MAX_POINTS = 16_000;
const MAX_SIZE = 20_000;
const COLOR = /^#[\da-f]{6}$/iu;
const NUMBER = String.raw`-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?`;
const PATH = new RegExp(
	String.raw`^M\s*${NUMBER}[ ,]+${NUMBER}(?:\s*L\s*${NUMBER}[ ,]+${NUMBER})*$`,
	'iu',
);
const TOOLS = new Set(['pen', 'pencil', 'highlighter', 'actionPen']);

const finite = (value: unknown, low: number, high: number): number | undefined =>
	typeof value === 'number' && Number.isFinite(value) && value >= low && value <= high
		? value
		: undefined;
const numbers = (value: unknown, count: number): number[] | undefined =>
	Array.isArray(value) &&
	value.length === count &&
	value.every((item) => typeof item === 'number' && Number.isFinite(item))
		? (value as number[])
		: undefined;

/** A pasted stroke rebuilt from what a stroke may hold, or undefined when it is not one. */
function cleanStroke(
	value: unknown,
	budget: { points: number },
): Record<string, unknown> | undefined {
	if (typeof value !== 'object' || value === null) return undefined;
	const stroke = value as Record<string, unknown>;
	const path = typeof stroke.path === 'string' ? stroke.path.trim() : '';
	const width = finite(stroke.width, Number.MIN_VALUE, 1000);
	const color =
		typeof stroke.color === 'string' && COLOR.test(stroke.color) ? stroke.color : undefined;
	if (!PATH.test(path) || width === undefined || !color) return undefined;
	const points = (path.match(/[ML]/giu) ?? []).length;
	budget.points -= points;
	if (budget.points < 0) return undefined;
	const clean: Record<string, unknown> = {
		path,
		color,
		width,
		opacity: finite(stroke.opacity, 0, 1) ?? 1,
	};
	if (typeof stroke.tool === 'string' && TOOLS.has(stroke.tool)) clean.tool = stroke.tool;
	for (const key of ['pressures', 'tiltAngles', 'tiltMagnitudes', 'pointTimestamps']) {
		const list = numbers(stroke[key], points);
		if (list) clean[key] = list;
	}
	for (const key of ['startTimestamp', 'durationMs']) {
		const time = finite(stroke[key], 0, Number.MAX_SAFE_INTEGER);
		if (time !== undefined) clean[key] = time;
	}
	return clean;
}

const H_FROM = new Set([
	'character',
	'column',
	'insideMargin',
	'leftMargin',
	'margin',
	'outsideMargin',
	'page',
	'rightMargin',
]);
const V_FROM = new Set([
	'bottomMargin',
	'insideMargin',
	'line',
	'margin',
	'outsideMargin',
	'page',
	'paragraph',
	'topMargin',
]);
const WRAPS = new Set(['square', 'tight', 'through', 'topAndBottom', 'none']);

/** A pasted floating position, kept to the values a drawing's placement may hold. */
function cleanPlacement(text: string | undefined): string | undefined {
	let value: unknown;
	try {
		value = text ? JSON.parse(text) : undefined;
	} catch {
		return undefined;
	}
	if (typeof value !== 'object' || value === null) return undefined;
	const placement = value as Record<string, unknown>;
	const wrap =
		typeof placement.wrap === 'string' && WRAPS.has(placement.wrap) ? placement.wrap : 'none';
	const offset = (key: string) => finite(placement[key], -MAX_SIZE * 10, MAX_SIZE * 10);
	const pick = (key: string, allowed: Set<string>) =>
		typeof placement[key] === 'string' && allowed.has(placement[key] as string)
			? placement[key]
			: undefined;
	const clean: Record<string, unknown> = {
		wrap,
		relativeFrom: pick('relativeFrom', H_FROM),
		relativeFromV: pick('relativeFromV', V_FROM),
		offsetXPx: offset('offsetXPx'),
		offsetYPx: offset('offsetYPx'),
		...(placement.behindText === true ? { behindText: true } : {}),
	};
	for (const key of Object.keys(clean)) if (clean[key] === undefined) delete clean[key];
	return JSON.stringify(clean);
}

/**
 * The ink node attributes read back from a pasted placeholder, rebuilt from what ink may hold
 * (pasted HTML can come from anywhere), or none when it holds no ink.
 */
export function inkFromClipboard(el: HTMLElement): Record<string, unknown> {
	const text = el.dataset.ink;
	if (!text || text.length > MAX_CLIPBOARD_INK) return {};
	let ink: { strokes?: unknown; width?: unknown; height?: unknown };
	try {
		ink = JSON.parse(text) as typeof ink;
	} catch {
		return {};
	}
	const width = finite(ink.width, Number.MIN_VALUE, MAX_SIZE);
	const height = finite(ink.height, Number.MIN_VALUE, MAX_SIZE);
	if (!Array.isArray(ink.strokes) || ink.strokes.length > MAX_STROKES || !width || !height)
		return {};
	const budget = { points: MAX_POINTS };
	const strokes = ink.strokes.map((stroke) => cleanStroke(stroke, budget));
	if (!strokes.length || strokes.some((stroke) => !stroke)) return {};
	const size = (key: string, fallback: number) =>
		finite(Number(el.dataset[key]), 1, MAX_SIZE) ?? fallback;
	const placement = el.dataset.anchored === '1' ? cleanPlacement(el.dataset.placement) : undefined;
	return {
		ink: JSON.stringify({ strokes, width, height }),
		widthPx: size('widthPx', width),
		heightPx: size('heightPx', height),
		anchored: placement !== undefined,
		...(placement ? { placement } : {}),
		drawingName: (el.dataset.drawingName || 'Ink').slice(0, 255),
	};
}
