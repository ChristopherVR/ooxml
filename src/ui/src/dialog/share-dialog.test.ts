import type { CollaborationState } from 'ooxml-core/collab';
import { OfficeUiShareDialog, registerControls, shareDialogLabels } from '../controls';
import { SHARE_DIALOG_LABELS } from './share-dialog-locales';

registerControls();

afterEach(() => {
	document.body.replaceChildren();
});

const create = (props: Partial<OfficeUiShareDialog> = {}): OfficeUiShareDialog => {
	const dialog = document.createElement('office-ui-share-dialog') as OfficeUiShareDialog;
	Object.assign(dialog, props);
	document.body.append(dialog);
	dialog.ensureRendered();
	return dialog;
};
const root = (dialog: Element) => dialog.shadowRoot as ShadowRoot;
const input = (dialog: Element, key: string) =>
	root(dialog).querySelector(`input[data-key="${key}"]`) as HTMLInputElement;
const type = (dialog: OfficeUiShareDialog, key: string, value: string) => {
	const field = input(dialog, key);
	field.value = value;
	field.dispatchEvent(new Event('input'));
	dialog.ensureRendered();
};
const submit = (dialog: Element) =>
	root(dialog).querySelector('[data-share="start"]') as HTMLButtonElement;

const live: CollaborationState = {
	active: true,
	roomId: 'team',
	status: 'connected',
	synced: true,
	people: [
		{ clientId: 1, name: 'Ann Lee', color: '#c43e1c', self: true },
		{ clientId: 2, name: 'Bo', color: '#107c41', self: false },
	],
};

describe('office-ui-share-dialog', () => {
	it('suggests a session name on open and starts once a name is typed', () => {
		const dialog = create({ open: true, roomPrefix: 'doc' });
		expect(input(dialog, 'roomId').value).toMatch(/^doc-[a-z0-9]{8}$/u);
		expect(submit(dialog).disabled).toBe(true);
		const started: unknown[] = [];
		dialog.addEventListener('office-share-start', (event) =>
			started.push((event as CustomEvent).detail.options),
		);
		type(dialog, 'userName', 'Ann');
		type(dialog, 'serverUrl', 'wss://s.example');
		expect(submit(dialog).disabled).toBe(false);
		submit(dialog).click();
		expect(started).toEqual([
			{
				roomId: input(dialog, 'roomId').value,
				serverUrl: 'wss://s.example',
				user: { name: 'Ann' },
				intent: 'create',
			},
		]);
	});

	it('flags an invalid session name and keeps Start disabled', () => {
		const dialog = create({ open: true, defaults: { userName: 'Ann' } });
		type(dialog, 'roomId', 'bad name');
		expect(root(dialog).querySelector('.error')?.textContent).toBe(
			SHARE_DIALOG_LABELS.en.invalidRoom,
		);
		expect(submit(dialog).disabled).toBe(true);
	});

	it('opens on the join tab with an invitation and joins it', () => {
		const dialog = create({
			open: true,
			invitation: 'https://x.example/?room=team&server=wss%3A%2F%2Fs',
			defaults: { userName: 'Bo' },
		});
		const tab = root(dialog).querySelector('[data-view="join"]');
		expect(tab?.getAttribute('aria-selected')).toBe('true');
		expect(dialog.formOptions()).toEqual({
			roomId: 'team',
			serverUrl: 'wss://s',
			user: { name: 'Bo' },
			intent: 'join',
		});
		expect(submit(dialog).textContent).toBe('Join');
	});

	it('shows the live session: link, details, people and Stop', () => {
		const dialog = create({
			open: true,
			state: live,
			defaults: { serverUrl: 'wss://s' },
			location: { origin: 'https://office.example', pathname: '/word/' },
		});
		const shadow = root(dialog);
		expect(shadow.querySelector('.link input')).toHaveProperty(
			'value',
			'https://office.example/word/?room=team&server=wss%3A%2F%2Fs',
		);
		expect(shadow.querySelector('.status')?.getAttribute('data-status')).toBe('connected');
		expect(
			Array.from(shadow.querySelectorAll('.people li .name')).map((node) => node.textContent),
		).toEqual(['Ann Lee', 'Bo']);
		expect(shadow.querySelector('.people .avatar')?.textContent).toBe('AL');
		let stopped = 0;
		dialog.addEventListener('office-share-stop', () => stopped++);
		(shadow.querySelector('[data-share="stop"]') as HTMLButtonElement).click();
		expect(stopped).toBe(1);
	});

	it('closes from Cancel with an event', () => {
		const dialog = create({ open: true });
		let closed = 0;
		dialog.addEventListener('office-share-close', () => closed++);
		(root(dialog).querySelector('button.cancel') as HTMLButtonElement).click();
		expect(closed).toBe(1);
		expect(dialog.open).toBe(false);
	});

	it('translates through the shipped locales', () => {
		const dialog = create({ open: true, labels: shareDialogLabels('de-AT') });
		expect(submit(dialog).textContent).toBe('Freigabe starten');
		expect(shareDialogLabels('zh-Hans').title).toBe(SHARE_DIALOG_LABELS['zh-CN'].title);
		expect(shareDialogLabels('pt')).toBe(SHARE_DIALOG_LABELS.en);
	});

	it('ships every caption in every locale', () => {
		const keys = Object.keys(SHARE_DIALOG_LABELS.en).sort();
		for (const labels of Object.values(SHARE_DIALOG_LABELS))
			expect(Object.keys(labels).sort()).toEqual(keys);
	});
});
