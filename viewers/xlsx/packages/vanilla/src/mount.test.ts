// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountEditor } from './index';

afterEach(() => document.body.replaceChildren());

describe('xlsx-vanilla-viewer mountEditor', () => {
	it('keeps in-editor changes across updates and reports them through the callbacks', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const onReadOnlyChange = vi.fn();
		const onRibbonCustomize = vi.fn();
		const options = { readOnly: false, locale: 'en', onReadOnlyChange, onRibbonCustomize };
		const binding = mountEditor(host, options);
		const element = binding.element;
		element.readOnly = true;
		element.locale = 'fr';
		expect(onReadOnlyChange).toHaveBeenCalledWith(true);
		element.dispatchEvent(
			new CustomEvent('ribbon-customize', { detail: { hiddenActions: ['bold'] }, bubbles: true }),
		);
		expect(onRibbonCustomize).toHaveBeenCalledWith(['bold']);
		binding.update({ ...options, authorName: 'Grace' });
		expect(element.readOnly).toBe(true);
		expect(element.locale).toBe('fr');
		expect(element.authorName).toBe('Grace');
		binding.destroy();
	});
});
