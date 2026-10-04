import {
	initialsOf,
	participantsFromAwareness,
	registerOfficeUi,
	type PresenceParticipant,
} from './index.js';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Presence = HTMLElement & { participants: PresenceParticipant[] };
const mount = (attrs = ''): Presence => {
	document.body.innerHTML = `<office-ui-presence ${attrs}></office-ui-presence>`;
	return document.body.firstElementChild as Presence;
};
const people: PresenceParticipant[] = [
	{ id: '1', name: 'Ada Lovelace', color: '#ff0000' },
	{ id: '2', name: 'Grace Hopper', status: 'idle' },
	{ id: '3', name: 'me', self: true },
	{ id: '4', name: 'Linus' },
	{ id: '5', name: 'Margaret Hamilton', status: 'away' },
];

describe('office-ui-presence', () => {
	it('renders initials, puts the local user first and names every avatar', () => {
		const el = mount('max="10"');
		el.participants = people;
		const buttons = [...el.shadowRoot!.querySelectorAll('button')];
		expect(buttons[0]!.getAttribute('aria-label')).toBe('me (you)');
		expect(buttons.map((b) => b.textContent)).toEqual(['M', 'AL', 'GH', 'L', 'MH']);
		expect(buttons[2]!.getAttribute('aria-label')).toBe('Grace Hopper, idle');
		expect(buttons[1]!.style.getPropertyValue('--office-avatar-color')).toBe('#ff0000');
		expect(el.shadowRoot!.querySelector('ul')!.getAttribute('role')).toBe('list');
		expect(el.getAttribute('role')).toBe('group');
	});

	it('collapses the overflow into a +N chip', () => {
		const el = mount('max="2"');
		el.participants = people;
		expect(el.shadowRoot!.querySelectorAll('button')).toHaveLength(2);
		const chip = el.shadowRoot!.querySelector('.more')!;
		expect(chip.textContent).toBe('+3');
		expect(chip.getAttribute('aria-label')).toContain('Linus');
		el.setAttribute('max', '5');
		expect(el.shadowRoot!.querySelector('.more')).toBeNull();
	});

	it('emits office-presence-select with the participant id', () => {
		const el = mount();
		el.participants = people;
		const seen: unknown[] = [];
		document.body.addEventListener('office-presence-select', (e) =>
			seen.push((e as CustomEvent).detail),
		);
		el.shadowRoot!.querySelector<HTMLButtonElement>('button[data-id="1"]')!.click();
		expect(seen).toEqual([{ id: '1' }]);
	});
});

describe('helpers', () => {
	it('initialsOf handles blanks, single and multi-word names', () => {
		expect(initialsOf('')).toBe('?');
		expect(initialsOf('ada')).toBe('A');
		expect(initialsOf('  grace   brewster hopper ')).toBe('GH');
	});

	it('participantsFromAwareness skips incomplete states and flags self', () => {
		const out = participantsFromAwareness(
			[
				{ clientId: 7, user: { name: 'Ada', color: '#00f' } },
				{ clientId: 8 },
				undefined,
				{ user: { id: 'u9', name: 'Me' } },
			],
			'u9',
		);
		expect(out).toEqual([
			{ id: '7', name: 'Ada', color: '#00f' },
			{ id: 'u9', name: 'Me', self: true },
		]);
	});
});
