// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createBuildingBlockSession } from './building-blocks';
import { navigationControls, statusControls, toolbarControls, type Control } from './controls';
import { createControlsView } from './controls-dom';
import { createDocument } from '../../core/src/index';

afterEach(() => document.body.replaceChildren());

function engine(canEdit = true) {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Several words in a row' }] }];
	const session = createBuildingBlockSession({ documentModel: model, canEdit });
	const host = document.createElement('div');
	document.body.append(host);
	session.canvasProps.mount(host);
	return session;
}
const flat = (groups: readonly (readonly Control[])[]) => groups.flat();
const byId = (groups: readonly (readonly Control[])[], id: string) =>
	flat(groups).find((control) => control.id === id);

describe('control descriptors', () => {
	it('cover history, font, formatting, alignment and lists with enabled and mixed state', () => {
		const session = engine();
		try {
			const groups = toolbarControls(session.state.toolbarProps!);
			expect(flat(groups).map((control) => control.id)).toEqual([
				'undo',
				'redo',
				'fontFamily',
				'fontSize',
				'bold',
				'italic',
				'underline',
				'alignLeft',
				'alignCenter',
				'alignRight',
				'alignJustify',
				'bulletList',
				'numberList',
			]);
			expect(byId(groups, 'undo')).toMatchObject({ disabled: true });
			expect(byId(groups, 'alignLeft')).toMatchObject({ pressed: true });
			expect(byId(groups, 'fontSize')).toMatchObject({ value: '11' });
			const bold = byId(groups, 'bold');
			expect(bold).toMatchObject({ kind: 'toggle', pressed: false, disabled: false });
			if (bold?.kind === 'toggle') bold.run();
			const after = toolbarControls(session.state.toolbarProps!);
			expect(byId(after, 'bold')).toMatchObject({ pressed: true });
			const numbered = byId(after, 'numberList');
			if (numbered?.kind === 'toggle') numbered.run();
			expect(byId(toolbarControls(session.state.toolbarProps!), 'numberList')).toMatchObject({
				pressed: true,
			});
			const center = byId(after, 'alignCenter');
			if (center?.kind === 'toggle') center.run();
			expect(byId(toolbarControls(session.state.toolbarProps!), 'alignCenter')).toMatchObject({
				pressed: true,
			});
			const size = byId(after, 'fontSize');
			if (size?.kind === 'select') size.change('20');
			expect(byId(toolbarControls(session.state.toolbarProps!), 'fontSize')).toMatchObject({
				value: '20',
			});
		} finally {
			session.destroy();
		}
	});

	it('disable every editing control read-only and list host workflows only when supplied', () => {
		const session = engine(false);
		try {
			const groups = toolbarControls(session.state.toolbarProps!);
			for (const id of ['bold', 'fontFamily', 'alignCenter', 'bulletList'])
				expect(byId(groups, id), id).toMatchObject({ disabled: true });
			expect(byId(groups, 'link')).toBeUndefined();
			expect(byId(groups, 'find')).toBeUndefined();
			session.update({
				canEdit: true,
				onOpenHyperlinkDialog: vi.fn(),
				onOpenFindReplace: vi.fn(),
				onOpenImagePicker: vi.fn(),
			});
			const supplied = toolbarControls(session.state.toolbarProps!);
			expect(['link', 'find', 'picture'].map((id) => Boolean(byId(supplied, id)))).toEqual([
				true,
				true,
				true,
			]);
		} finally {
			session.destroy();
		}
	});

	it('describe page navigation and status from their props', () => {
		const session = engine();
		try {
			const navigation = navigationControls(session.state.navigationProps!);
			expect(byId(navigation, 'pageText')).toMatchObject({ text: 'Page 1 of 1' });
			expect(byId(navigation, 'previousPage')).toMatchObject({ disabled: true });
			expect(byId(navigation, 'viewMode')).toMatchObject({ value: 'draft' });
			const status = statusControls(session.state.statusProps!);
			expect(byId(status, 'wordText')).toMatchObject({ text: '5 words' });
			const zoom = byId(status, 'zoom');
			expect(zoom).toMatchObject({ value: '100' });
			if (zoom?.kind === 'select') zoom.change('150');
			expect(session.state.statusProps?.zoom).toBe(150);
		} finally {
			session.destroy();
		}
	});
});

describe('DOM controls view', () => {
	it('updates nodes in place so focus survives and removes controls that disappear', () => {
		const session = engine();
		const host = document.createElement('div');
		document.body.append(host);
		const view = createControlsView(host);
		try {
			view.update(toolbarControls(session.state.toolbarProps!));
			const bold = host.querySelector<HTMLButtonElement>('[data-control="bold"]')!;
			bold.focus();
			expect(bold.getAttribute('aria-pressed')).toBe('false');
			bold.click();
			view.update(toolbarControls(session.state.toolbarProps!));
			expect(host.querySelector('[data-control="bold"]')).toBe(bold);
			expect(bold.getAttribute('aria-pressed')).toBe('true');
			expect(document.activeElement).toBe(bold);
			const family = host.querySelector<HTMLSelectElement>('[data-control="fontFamily"]')!;
			expect(family.value).toBe('Calibri');
			family.value = 'Arial';
			family.dispatchEvent(new Event('change'));
			view.update(toolbarControls(session.state.toolbarProps!));
			expect(host.querySelector('[data-control="fontFamily"]')).toBe(family);
			expect(family.value).toBe('Arial');
			session.update({ canEdit: true, onOpenFindReplace: vi.fn() });
			view.update(toolbarControls(session.state.toolbarProps!));
			expect(host.querySelector('[data-control="find"]')).not.toBeNull();
			session.update({ canEdit: true });
			view.update(toolbarControls(session.state.toolbarProps!));
			expect(host.querySelector('[data-control="find"]')).toBeNull();
			view.destroy();
			expect(host.childElementCount).toBe(0);
		} finally {
			session.destroy();
		}
	});
});
