import { TEAMS_MESSAGES } from '../i18n/translator';
import { defineTeamsApp, type TeamsApp } from './teams-app';

const LOCAL = JSON.stringify({ mode: 'local', iceServers: [] });
const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));
const q = <T extends Element>(root: ParentNode | null | undefined, s: string): T =>
	root!.querySelector<T>(s)!;

beforeAll(() => defineTeamsApp());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

let workspace = 0;
async function mount(locale?: string): Promise<TeamsApp> {
	const el = document.createElement('teams-app') as TeamsApp;
	el.setAttribute('workspace-id', `locale-${Date.now()}-${workspace++}`);
	el.setAttribute('user-name', 'Ada');
	el.setAttribute('user-id', 'ada');
	el.setAttribute('server-config', LOCAL);
	if (locale) el.setAttribute('locale', locale);
	document.body.append(el);
	await el.updateComplete;
	await settle(el);
	return el;
}

/** Let the app and every Teams element under it finish rendering. */
async function settle(root: Element): Promise<void> {
	for (let round = 0; round < 4; round++) {
		await tick(10);
		for (const el of everyElement(root))
			await (el as { updateComplete?: Promise<unknown> }).updateComplete;
	}
}

function* everyElement(root: ParentNode): Generator<Element> {
	for (const el of root.querySelectorAll('*')) {
		yield el;
		if (el.shadowRoot) yield* everyElement(el.shadowRoot);
	}
}

const ATTRIBUTES = ['aria-label', 'title', 'placeholder', 'label', 'alt'];

/** Every piece of text a person could read or hear, in the whole shadow tree. */
function visibleStrings(root: Element): string[] {
	const out: string[] = [];
	const add = (text: string | null | undefined): void => {
		const clean = (text ?? '').replace(/\s+/gu, ' ').trim();
		if (clean) out.push(clean);
	};
	const walk = (parent: ParentNode): void => {
		for (const node of parent.childNodes) {
			if (node.nodeType === Node.TEXT_NODE) add(node.textContent);
			else if (node instanceof Element) {
				if (['STYLE', 'SCRIPT'].includes(node.tagName)) continue;
				for (const name of ATTRIBUTES) add(node.getAttribute(name));
				walk(node);
				if (node.shadowRoot) walk(node.shadowRoot);
			}
		}
	};
	walk(root);
	if (root.shadowRoot) walk(root.shadowRoot);
	return out;
}

/** Drive the app through its screens and gather the text of each. */
async function sweep(el: TeamsApp): Promise<string[]> {
	const seen = new Set<string>();
	const take = async (): Promise<void> => {
		await settle(el);
		for (const text of visibleStrings(el)) seen.add(text);
	};
	el.creatingChannel = true;
	el.newChat = true;
	el.meetingForm = { meetingId: '', start: 0 };
	el.addingTab = true;
	await take();
	el.addingTab = false;
	el.creatingChannel = false;
	el.newChat = false;
	el.meetingForm = null;
	for (const rail of [
		'activity',
		'chat',
		'calendar',
		'calls',
		'files',
		'followed',
		'drafts',
		'teams',
	] as const) {
		el.rail = rail;
		await take();
	}
	el.settingsOpen = true;
	const settings = el.shadowRoot!.querySelector('teams-settings') as HTMLElement & {
		category: string;
	};
	for (const category of ['general', 'appearance', 'notifications', 'files', 'connection']) {
		settings.category = category;
		await take();
	}
	settings.category = 'general';
	(settings as unknown as { query: string }).query = 'theme';
	await take();
	return [...seen];
}

describe('<teams-app locale>', () => {
	it('renders German with no English wording left', async () => {
		const el = await mount('de');
		const german = TEAMS_MESSAGES.de;
		const english = TEAMS_MESSAGES.en;
		const strings = await sweep(el);
		expect(strings.length).toBeGreaterThan(150);

		// Positive: German is really on screen.
		const text = strings.join('\n');
		for (const word of ['Kanäle', 'Aktivität', 'Einstellungen', 'Kalender', 'Anrufe'])
			expect(text, word).toContain(word);

		// Negative 1: no static English message appears where its German twin should be.
		const leaked = Object.keys(english).filter((key) => {
			const en = english[key]!;
			if (/\{\w+\}/u.test(en) || en === german[key] || en.length < 4) return false;
			return strings.some((s) => s === en || (en.length >= 12 && s.includes(en)));
		});
		expect(leaked).toEqual([]);

		// Negative 2: no common English function word survives in any visible string.
		const english2 = /\b(the|your|this|with|from|and|not|yet|you|please|click|choose)\b/u;
		const offenders = strings.filter((s) => english2.test(s));
		expect(offenders).toEqual([]);
	});

	it('translates the rail, the toolbar and the composer placeholder', async () => {
		const el = await mount('de');
		const rail = q<HTMLElement>(el.shadowRoot, 'office-ui-app-rail');
		expect(rail.getAttribute('aria-label')).toBe('App-Leiste');
		const labels = [...rail.shadowRoot!.querySelectorAll('ul button .label')].map((n) =>
			n.textContent?.trim(),
		);
		expect(labels).toEqual(expect.arrayContaining(['Aktivität', 'Kalender', 'Chat']));
		const composer = q<HTMLElement>(el.shadowRoot, 'office-ui-chat-composer');
		expect(composer.shadowRoot!.querySelector('textarea')!.getAttribute('aria-label')).toBe(
			'Nachricht',
		);
	});

	it('follows a locale change at runtime, down to the nested elements', async () => {
		const el = await mount('en');
		const rail = q<HTMLElement>(el.shadowRoot, 'office-ui-app-rail');
		expect(rail.getAttribute('aria-label')).toBe('App bar');
		el.locale = 'fr';
		await settle(el);
		expect(rail.getAttribute('aria-label')).toBe("Barre d'applications");
		el.setAttribute('locale', 'zh-CN');
		await settle(el);
		expect(rail.getAttribute('aria-label')).toBe('应用栏');
		el.locale = '';
		await settle(el);
		expect(rail.getAttribute('aria-label')).toBe('App bar');
	});

	it('keeps the locale of a standalone element that was given its own', async () => {
		const el = await mount('de');
		const list = document.createElement('office-ui-chat-list');
		list.setAttribute('locale', 'es');
		el.shadowRoot!.append(list);
		await settle(el);
		expect(list.getAttribute('aria-label')).toBe('Mensajes');
		el.locale = 'fr';
		await settle(el);
		// The host pushes its locale to what it rendered, but it never overrides a parent-less lone element.
		expect(q<HTMLElement>(el.shadowRoot, 'office-ui-app-rail').getAttribute('aria-label')).toBe(
			"Barre d'applications",
		);
	});

	it('formats dates and times with Intl in the display locale', async () => {
		const el = await mount('de');
		const calendar = document.createElement('office-ui-calendar-view') as HTMLElement & {
			now: number;
			zone: string;
			locale: string;
		};
		calendar.locale = 'de';
		calendar.zone = 'UTC';
		calendar.now = Date.UTC(2026, 9, 15, 12);
		el.shadowRoot!.append(calendar);
		await settle(calendar);
		const range = calendar.shadowRoot!.querySelector('.range')!.textContent!;
		expect(range).toMatch(/Oktober/u);
		const heads = [...calendar.shadowRoot!.querySelectorAll('.dow')].map((n) => n.textContent);
		expect(heads.join(' ')).toMatch(/Mo\.?|Di\.?|Mi\.?/u);
	});

	it('words the notices of the core in the display locale', async () => {
		const el = await mount('de');
		el.client!.createGroupChat([], '');
		await settle(el);
		expect(el.shadowRoot!.querySelector('.toast')?.textContent).toBe(
			'Wählen Sie mindestens eine Person aus',
		);
	});
});
