import { afterEach, describe, expect, it } from 'vitest';
import { mountViewer } from './binding';
import { demoDocument } from 'ooxml-core/visio/ui';

afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

type Profile = { displayName: string; avatarColor: string; initial?: string };

function setup() {
	localStorage.setItem(
		'ooxml-office-profile',
		JSON.stringify({ displayName: 'Ada Lovelace', avatarColor: '#16a34a' }),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: structuredClone(demoDocument) });
	const root = viewer.element.shadowRoot!;
	const backstage = root.querySelector<HTMLElement & { open: boolean }>('office-ui-backstage')!;
	const item = (id: string) =>
		backstage.shadowRoot!.querySelector<HTMLButtonElement>(`[data-backstage-item="${id}"]`)!;
	const file = () =>
		root.querySelector('office-ui-ribbon')!.shadowRoot!.querySelector<HTMLButtonElement>('.file')!;
	const dialog = root.querySelector<HTMLElement & { open: boolean }>('office-ui-options-dialog')!;
	const account = root.querySelector<HTMLElement & { profile: Profile }>('office-ui-account')!;
	const field = (key: string) =>
		dialog.shadowRoot!.querySelector<HTMLInputElement>(
			`[data-key="${key}"] input, [data-key="${key}"] select`,
		)!;
	return { viewer, root, item, file, backstage, dialog, account, field };
}

describe('Visio File > Account and Options', () => {
	it('shows the shared profile on the Account page', () => {
		const { viewer, root, item, file, account } = setup();
		file().click();
		item('account').click();
		const page = root.querySelector<HTMLElement>('[data-backstage-page="account"]')!;
		expect(page.hidden).toBe(false);
		expect(page.contains(account)).toBe(true);
		expect(account.profile).toEqual({ displayName: 'Ada Lovelace', avatarColor: '#16a34a' });
		expect(page.querySelector('.backstage-product h2')!.textContent).toBe('Product Information');
		viewer.destroy();
	});

	it('opens Visio Options over the drawing with its categories in Visio order', () => {
		const { viewer, item, file, backstage, dialog, field } = setup();
		file().click();
		item('options').click();
		expect(backstage.open).toBe(false);
		expect(dialog.open).toBe(true);
		expect(dialog.getAttribute('heading')).toBe('Visio Options');
		const tabs = [...dialog.shadowRoot!.querySelectorAll('[role="tab"]')].map(
			(tab) => tab.textContent,
		);
		expect(tabs).toEqual([
			'General',
			'Proofing',
			'Save',
			'Language',
			'Advanced',
			'Customize Ribbon',
			'Quick Access Toolbar',
			'Add-ins',
			'Trust Center',
		]);
		// What the viewer honours is enabled; what it cannot do yet says why.
		expect(field('livePreview').disabled).toBe(false);
		expect(field('miniToolbar').disabled).toBe(false);
		expect(field('userName').value).toBe('Ada Lovelace');
		(dialog as unknown as { category: string }).category = 'add-ins';
		expect(dialog.shadowRoot!.querySelector('.unavailable')!.textContent).toMatch(/^Not available/);
		viewer.destroy();
	});

	it('writes Options edits to the profile and mirrors Account edits back', () => {
		const { viewer, item, file, dialog, account, field } = setup();
		file().click();
		item('options').click();
		const name = field('userName');
		name.value = 'Grace Hopper';
		name.dispatchEvent(new Event('input'));
		const initials = field('userInitials');
		initials.value = 'GH';
		initials.dispatchEvent(new Event('input'));
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
		expect(dialog.open).toBe(false);
		expect(account.profile).toEqual({
			displayName: 'Grace Hopper',
			avatarColor: '#16a34a',
			initial: 'GH',
		});
		expect(JSON.parse(localStorage.getItem('ooxml-office-profile')!).displayName).toBe(
			'Grace Hopper',
		);
		account.shadowRoot!.querySelector<HTMLButtonElement>('[data-color="#c2431f"]')!.click();
		expect(JSON.parse(localStorage.getItem('ooxml-office-profile')!).avatarColor).toBe('#c2431f');
		viewer.destroy();
	});

	it('applies the undo depth and ScreenTip style, and remembers them', () => {
		const { viewer, root, item, file, dialog, field } = setup();
		file().click();
		item('options').click();
		const tips = field('screenTips') as unknown as HTMLSelectElement;
		expect(tips.closest('label')!.title).not.toMatch(/^Not available yet/);
		tips.value = 'off';
		tips.dispatchEvent(new Event('change'));
		(dialog as unknown as { category: string }).category = 'advanced';
		const undo = field('undoLevels');
		undo.value = '5';
		undo.dispatchEvent(new Event('change'));
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
		expect(viewer.element.controller.undoLevels).toBe(5);
		expect(
			root.querySelectorAll('office-ui-ribbon [title]:not([title*="not available"])'),
		).toHaveLength(0);
		expect(JSON.parse(localStorage.getItem('ooxml-visio-options')!)).toMatchObject({
			screenTips: 'off',
			undoLevels: 5,
		});
		viewer.destroy();
	});

	it('switches the status bar between the neutral and Colorful Office themes and remembers it', () => {
		const { viewer, item, file, dialog, field } = setup();
		// Neutral by default; a host may set the attribute or property itself.
		expect(viewer.element.statusBar).toBe('neutral');
		file().click();
		item('options').click();
		const theme = field('officeTheme') as unknown as HTMLSelectElement;
		expect(theme.value).toBe('neutral');
		theme.value = 'colorful';
		theme.dispatchEvent(new Event('change'));
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
		expect(viewer.element.getAttribute('status-bar')).toBe('colorful');
		expect(viewer.element.statusBar).toBe('colorful');
		expect(localStorage.getItem('ooxml-visio-office-theme')).toBe('colorful');
		viewer.destroy();
		// A new viewer starts with the saved choice, unless its host chose one.
		const next = mountViewer(document.body.appendChild(document.createElement('div')));
		expect(next.element.statusBar).toBe('colorful');
		next.destroy();
		const chosen = document.createElement('visio-viewer');
		chosen.setAttribute('status-bar', 'neutral');
		document.body.append(chosen);
		expect(chosen.statusBar).toBe('neutral');
		chosen.remove();
	});

	describe('behaviour options', () => {
		type Ctx = ReturnType<typeof setup>;
		/** Open Options on `category`, set `key` through its control, and press OK. */
		function change(ctx: Ctx, category: string, key: string, value: string | boolean): void {
			ctx.file().click();
			ctx.item('options').click();
			(ctx.dialog as unknown as { category: string }).category = category;
			const control = ctx.field(key) as unknown as HTMLInputElement | HTMLSelectElement;
			if (typeof value === 'boolean') (control as HTMLInputElement).checked = value;
			else control.value = value;
			control.dispatchEvent(new Event('change'));
			ctx.dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
		}
		const ribbon = (ctx: Ctx) =>
			ctx.root.querySelector('office-ui-ribbon') as HTMLElement & { selected: string };
		const tabHidden = (ctx: Ctx, tab: string) =>
			ribbon(ctx).querySelector(`[data-ribbon-tab="${tab}"]`)!.hasAttribute('data-tab-hidden');
		const stored = () => JSON.parse(localStorage.getItem('ooxml-visio-options') ?? '{}');

		it('hides a ribbon tab and remembers it', () => {
			const ctx = setup();
			expect(tabHidden(ctx, 'insert')).toBe(false);
			change(ctx, 'ribbon', 'tab-insert', false);
			expect(tabHidden(ctx, 'insert')).toBe(true);
			expect(stored()['tab-insert']).toBe(false);
			ctx.viewer.destroy();
		});

		it('moves off a tab that is hidden while it is open', () => {
			const ctx = setup();
			ribbon(ctx).selected = 'view';
			change(ctx, 'ribbon', 'tab-view', false);
			expect(ribbon(ctx).selected).not.toBe('view');
			ctx.viewer.destroy();
		});

		it('shows only the Quick Access Toolbar commands that are switched on', () => {
			const ctx = setup();
			const bar = () =>
				ctx.root.querySelector('office-ui-title-bar') as HTMLElement & {
					state: { quickAccess: { items: { id: string }[] } };
				};
			const ids = () => bar().state.quickAccess.items.map((item) => item.id);
			expect(ids()).toEqual(['save', 'undo', 'redo']);
			change(ctx, 'qat', 'qat-redo', false);
			expect(ids()).toEqual(['save', 'undo']);
			ctx.viewer.destroy();
		});

		it('translates the ribbon into the chosen display language, unless the host chose one', () => {
			const ctx = setup();
			change(ctx, 'language', 'displayLanguage', 'de');
			expect(ribbon(ctx).dataset.visioLocale).toBe('de');
			expect(stored().displayLanguage).toBe('de');
			ctx.viewer.destroy();
			const host = document.createElement('div');
			document.body.append(host);
			const fixed = mountViewer(host);
			fixed.element.setAttribute('locale', 'fr');
			const fixedRibbon = fixed.element.shadowRoot!.querySelector(
				'office-ui-ribbon',
			) as HTMLElement;
			expect(fixedRibbon.dataset.visioLocale).toBe('fr');
			fixed.destroy();
		});

		it('turns the browser spell checker on or off for the text editor', () => {
			const ctx = setup();
			change(ctx, 'proofing', 'checkSpelling', false);
			expect(stored().checkSpelling).toBe(false);
			ctx.viewer.destroy();
		});

		it('keeps AutoRecover off until it is chosen, and remembers the choice and interval', () => {
			const ctx = setup();
			expect(stored().autoRecover).toBeUndefined();
			change(ctx, 'save', 'autoRecover', true);
			expect(stored()).toMatchObject({ autoRecover: true, autoRecoverMinutes: 10 });
			change(ctx, 'save', 'autoRecoverMinutes', '3');
			expect(stored().autoRecoverMinutes).toBe(3);
			ctx.viewer.destroy();
		});

		it('keeps one tab when every tab is switched off in storage', () => {
			const ids = ['home', 'insert', 'design', 'data', 'process', 'review', 'view', 'help'];
			localStorage.setItem(
				'ooxml-visio-options',
				JSON.stringify(Object.fromEntries(ids.map((id) => [`tab-${id}`, false]))),
			);
			const ctx = setup();
			expect(tabHidden(ctx, 'home')).toBe(false);
			expect(ids.filter((id) => tabHidden(ctx, id))).toHaveLength(7);
			ctx.viewer.destroy();
		});
	});
});
