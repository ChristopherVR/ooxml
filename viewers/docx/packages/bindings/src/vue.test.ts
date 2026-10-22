// @vitest-environment jsdom
import { createApp, h } from 'vue';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { WordEditor } from './vue';

afterEach(() => document.body.replaceChildren());

describe('Vue Word adapter: readOnly', () => {
	it('update:readOnly mirrors the host switch only, never a collaboration role lock', () => {
		const container = document.createElement('div');
		document.body.append(container);
		const onUpdate = vi.fn();
		const onChange = vi.fn();
		const app = createApp({
			render: () =>
				h(WordEditor, {
					readOnly: false,
					'onReadonly-change': onChange,
					'onUpdate:readOnly': onUpdate,
				}),
		});
		app.mount(container);
		const editor = container.querySelector('docx-editor')!;
		const send = (detail: object) =>
			editor.dispatchEvent(new CustomEvent('readonly-change', { detail }));
		send({ readOnly: true, locked: false, hostReadOnly: true });
		expect(onUpdate).toHaveBeenCalledWith(true);
		onUpdate.mockClear();
		send({ readOnly: true, locked: true, hostReadOnly: false });
		expect(onUpdate).not.toHaveBeenCalled();
		expect(onChange).toHaveBeenLastCalledWith(true, {
			readOnly: true,
			locked: true,
			hostReadOnly: false,
		});
		app.unmount();
	});
});
