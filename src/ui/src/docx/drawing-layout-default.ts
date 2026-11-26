import type { DrawingWrap } from 'ooxml-core/docx';

const key = 'ooxml-word-drawing-wrap';
const values = new Set<string>([
	'inline',
	'square',
	'tight',
	'through',
	'topAndBottom',
	'behind',
	'front',
]);
const session = new WeakMap<Window, Map<string, DrawingWrap>>();

export function isDrawingWrap(value: string): value is DrawingWrap {
	return values.has(value);
}

export function defaultDrawingWrap(
	doc: Document,
	kind: 'picture' | 'shape',
): DrawingWrap | undefined {
	const win = doc.defaultView;
	if (!win) return;
	try {
		const saved = win.localStorage.getItem(`${key}.${kind}`);
		if (saved && isDrawingWrap(saved)) return saved;
	} catch {
		return session.get(win)?.get(kind);
	}
}

export function saveDrawingWrap(doc: Document, kind: 'picture' | 'shape', wrap: DrawingWrap): void {
	const win = doc.defaultView;
	if (!win) return;
	try {
		win.localStorage.setItem(`${key}.${kind}`, wrap);
	} catch {
		const defaults = session.get(win) ?? new Map();
		defaults.set(kind, wrap);
		session.set(win, defaults);
	}
}
