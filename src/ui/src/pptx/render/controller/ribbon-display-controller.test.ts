import { describe, expect, it, vi } from 'vitest';

import { createRibbonDisplayController } from './ribbon-display-controller';

const key = (name: string, ctrlKey = false) => ({ key: name, ctrlKey, preventDefault: vi.fn() });

describe('ribbon display controller', () => {
	it('starts from the host value and reports only what the user picks', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ initial: 'tabs', onChange });
		expect(controller.getSnapshot().mode).toBe('tabs');

		controller.setMode('fullScreen');
		expect(controller.getSnapshot().layout.reveal).toBe(true);
		expect(onChange).not.toHaveBeenCalled();

		controller.choose('always');
		expect(onChange).toHaveBeenCalledExactlyOnceWith('always');
		expect(controller.getSnapshot().menu.find((item) => item.checked)?.mode).toBe('always');
		controller.choose('always');
		expect(onChange).toHaveBeenCalledOnce();
	});

	it('peeks the commands for a chosen tab in tabs mode, never for File, and puts them away', () => {
		const controller = createRibbonDisplayController({ initial: 'tabs' });
		controller.tabChosen('file');
		expect(controller.getSnapshot().peek).toBe(false);
		controller.tabChosen('insert');
		expect(controller.getSnapshot().layout.commands).toBe('overlay');
		controller.dismiss();
		expect(controller.getSnapshot().layout.commands).toBe('hidden');

		// A docked ribbon has nothing to peek.
		controller.setMode('always');
		controller.tabChosen('insert');
		expect(controller.getSnapshot().peek).toBe(false);
	});

	it('brings the ribbon back from the strip in full-screen mode only', () => {
		const controller = createRibbonDisplayController({ initial: 'fullScreen' });
		controller.reveal();
		expect(controller.getSnapshot().layout).toMatchObject({ tabs: true, reveal: false });
		controller.setMode('tabs');
		controller.reveal();
		expect(controller.getSnapshot().peek).toBe(false);
	});

	it('handles Ctrl+F1 and Escape, and ignores other keys', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ onChange });
		const toggle = key('F1', true);
		expect(controller.handleKeydown(toggle)).toBe(true);
		expect(toggle.preventDefault).toHaveBeenCalledOnce();
		expect(controller.getSnapshot().mode).toBe('tabs');
		expect(onChange).toHaveBeenLastCalledWith('tabs');

		expect(controller.handleKeydown(key('Escape'))).toBe(false);
		controller.toggleMenu();
		expect(controller.getSnapshot().menuOpen).toBe(true);
		expect(controller.handleKeydown(key('Escape'))).toBe(true);
		expect(controller.getSnapshot().menuOpen).toBe(false);
		expect(controller.handleKeydown(key('F1'))).toBe(false);
		expect(controller.handleKeydown(key('a', true))).toBe(false);
	});

	it('collapses automatically on a short viewport until the user chooses for themselves', () => {
		const onChange = vi.fn();
		const controller = createRibbonDisplayController({ onChange });
		controller.setShortViewport(true);
		expect(controller.getSnapshot().effectiveMode).toBe('always');

		controller.setAutoCollapse(true);
		expect(controller.getSnapshot()).toMatchObject({ mode: 'always', effectiveMode: 'tabs' });
		expect(controller.getSnapshot().menu.find((item) => item.checked)?.mode).toBe('tabs');
		expect(onChange).not.toHaveBeenCalled();

		// Ctrl+F1 from the collapsed state docks the ribbon and stops the automatic collapse.
		controller.toggleCollapsed();
		expect(controller.getSnapshot()).toMatchObject({ mode: 'always', effectiveMode: 'always' });
		expect(onChange).toHaveBeenLastCalledWith('always');

		// Switching the setting off and on again is a fresh instruction and applies once more.
		controller.setAutoCollapse(false);
		controller.setAutoCollapse(true);
		expect(controller.getSnapshot().effectiveMode).toBe('tabs');

		const taller = createRibbonDisplayController();
		taller.setAutoCollapse(true);
		expect(taller.getSnapshot().effectiveMode).toBe('always');
	});

	it('notifies subscribers once per change and keeps the snapshot when nothing changed', () => {
		const controller = createRibbonDisplayController();
		const listener = vi.fn();
		const stop = controller.subscribe(listener);
		const before = controller.getSnapshot();
		controller.dismiss();
		controller.setMode('always');
		expect(controller.getSnapshot()).toBe(before);
		expect(listener).not.toHaveBeenCalled();
		controller.setMode('tabs');
		expect(listener).toHaveBeenCalledOnce();
		stop();
		controller.setMode('always');
		expect(listener).toHaveBeenCalledOnce();
	});
});
