import { describe, expect, it, vi } from 'vitest';

import type { EditorKeyDispatchEvent, EditorKeyHandlers } from './editor-key-dispatch';
import { dispatchEditorKey } from './editor-key-dispatch';

function key(
	k: string,
	mods: Partial<Pick<EditorKeyDispatchEvent, 'ctrlKey' | 'shiftKey' | 'altKey' | 'metaKey'>> = {},
	target: EventTarget | null = null,
): EditorKeyDispatchEvent & { preventDefault: ReturnType<typeof vi.fn> } {
	return { key: k, ...mods, target, preventDefault: vi.fn() };
}

describe('dispatchEditorKey', () => {
	it('calls the handler named after the resolved action and claims the key', () => {
		const undo = vi.fn();
		const event = key('z', { ctrlKey: true });
		expect(dispatchEditorKey(event, {}, { undo })).toBe('undo');
		expect(undo).toHaveBeenCalledOnce();
		expect(event.preventDefault).toHaveBeenCalledOnce();
	});

	it('passes the nudge delta through', () => {
		const nudge = vi.fn();
		dispatchEditorKey(key('ArrowRight', { shiftKey: true }), { hasSelection: true }, { nudge });
		expect(nudge).toHaveBeenCalledWith(10, 0);
	});

	it('claims an editor chord even without a handler, unless it is passed through', () => {
		const claimed = key('ArrowLeft');
		expect(dispatchEditorKey(claimed, {}, {})).toBe('prevSlide');
		expect(claimed.preventDefault).toHaveBeenCalled();

		const prevSlide = vi.fn();
		const left = key('ArrowLeft');
		expect(dispatchEditorKey(left, { passThrough: ['prevSlide'] }, { prevSlide })).toBeNull();
		expect(prevSlide).not.toHaveBeenCalled();
		expect(left.preventDefault).not.toHaveBeenCalled();
	});

	it('leaves keys alone that map to nothing or are gated', () => {
		const handlers: EditorKeyHandlers = { delete: vi.fn() };
		const event = key('Delete');
		expect(dispatchEditorKey(event, { hasSelection: false }, handlers)).toBeNull();
		expect(
			dispatchEditorKey(key('Delete'), { hasSelection: true, canEdit: false }, handlers),
		).toBeNull();
		expect(event.preventDefault).not.toHaveBeenCalled();
		expect(handlers.delete).not.toHaveBeenCalled();
	});

	it('reads the text-input gate from the event target when the state omits it', () => {
		const input = { tagName: 'INPUT' } as unknown as EventTarget;
		const copy = vi.fn();
		expect(
			dispatchEditorKey(key('c', { ctrlKey: true }, input), { hasSelection: true }, { copy }),
		).toBeNull();
		expect(copy).not.toHaveBeenCalled();
	});

	it('starts the show on F5 ahead of the editor gates, when wired and allowed', () => {
		const startShowFromBeginning = vi.fn();
		const startShowFromCurrent = vi.fn();
		const handlers = { startShowFromBeginning, startShowFromCurrent };
		expect(dispatchEditorKey(key('F5'), { canEdit: false }, handlers)).toBe(
			'startShowFromBeginning',
		);
		expect(dispatchEditorKey(key('F5', { shiftKey: true }), {}, handlers)).toBe(
			'startShowFromCurrent',
		);
		expect(dispatchEditorKey(key('F5'), { canStartShow: false }, handlers)).toBeNull();
		expect(dispatchEditorKey(key('F5'), { isShowRunning: true }, handlers)).toBeNull();
		expect(startShowFromBeginning).toHaveBeenCalledOnce();
		expect(startShowFromCurrent).toHaveBeenCalledOnce();
	});

	it('honours disabled commands from the keyboard customisation', () => {
		const undo = vi.fn();
		const keyboard = { disableAll: false, disabled: new Set(['undo'] as const), remap: new Map() };
		expect(dispatchEditorKey(key('z', { ctrlKey: true }), { keyboard }, { undo })).toBeNull();
		expect(undo).not.toHaveBeenCalled();
	});
});
