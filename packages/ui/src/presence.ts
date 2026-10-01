import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/**
 * One collaborator as the avatar stack needs it. This is a deliberately minimal structural
 * view of an awareness entry (identity, colour, activity); core's `collab` awareness model
 * satisfies it, and `participantsFromAwareness` adapts raw `{ user }` awareness states.
 */
export interface PresenceParticipant {
	id: string;
	name: string;
	/** CSS colour of the participant's cursor and avatar ring. */
	color?: string;
	status?: 'active' | 'idle' | 'away';
	/** The local user: sorted first and labelled "(you)". */
	self?: boolean;
}

export type OfficePresenceSelectEvent = CustomEvent<{ id: string }>;

/** Adapt awareness-like states (`{ clientId?, user: { id?, name, color? } }`) to participants. */
export function participantsFromAwareness(
	states: Iterable<{ clientId?: string | number; user?: Partial<PresenceParticipant> } | undefined>,
	selfId?: string,
): PresenceParticipant[] {
	const out: PresenceParticipant[] = [];
	for (const state of states) {
		const user = state?.user;
		const id = String(user?.id ?? state?.clientId ?? '');
		if (!id || !user?.name) continue;
		out.push({
			id,
			name: user.name,
			...(user.color ? { color: user.color } : {}),
			...(user.status ? { status: user.status } : {}),
			...(id === selfId || user.self ? { self: true } : {}),
		});
	}
	return out;
}

export const initialsOf = (name: string): string => {
	const parts = name.trim().split(/\s+/).filter(Boolean);
	if (parts.length === 0) return '?';
	const first = [...(parts[0] ?? '')][0] ?? '?';
	const last = parts.length > 1 ? ([...(parts[parts.length - 1] ?? '')][0] ?? '') : '';
	return (first + last).toUpperCase();
};

const CSS = `
:host { display: inline-flex; align-items: center; }
ul { display: flex; margin: 0; padding: 0; list-style: none; }
li { margin-inline-start: -6px; }
li:first-child { margin-inline-start: 0; }
button, .more { box-sizing: border-box; display: inline-grid; place-items: center;
	width: var(--office-target-size, 28px); height: var(--office-target-size, 28px); padding: 0; border-radius: 50%;
	border: 2px solid var(--office-avatar-color, #6b7280); background: var(--office-surface, #f3f4f6);
	color: var(--office-foreground, #1f2937); font: 600 11px var(--office-font, system-ui, sans-serif); cursor: pointer; }
.more { cursor: default; border-color: var(--office-border, #d1d5db); }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 1px; z-index: 1; }
button[data-status="idle"] { opacity: .7; }
button[data-status="away"] { opacity: .45; border-style: dashed; }
@media (forced-colors: active) {
	button, .more { border-color: CanvasText; background: Canvas; color: CanvasText; }
	button:focus-visible { outline-color: Highlight; }
}
`;

/**
 * Presence avatar stack. `participants` property (array of `PresenceParticipant`), `max`
 * attribute (visible avatars, default 4; the rest collapse into a "+N" chip). `role="list"`
 * with one button per person (accessible name = name, status and "(you)"); activating one emits
 * `office-presence-select` `{ id }` so the host can follow that user. Rendering carries no
 * identity beyond what the host passes in.
 */
export const definePresence = definer('office-ui-presence', () => {
	class OfficeUiPresence extends HTMLElement {
		static observedAttributes = ['max', 'label'];
		private people: PresenceParticipant[] = [];
		private readonly list: HTMLUListElement;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			this.list = this.ownerDocument.createElement('ul');
			this.list.setAttribute('role', 'list');
			root.append(this.list);
		}
		connectedCallback(): void {
			this.render();
		}
		attributeChangedCallback(): void {
			this.render();
		}
		get participants(): PresenceParticipant[] {
			return this.people;
		}
		set participants(next: PresenceParticipant[]) {
			this.people = [...next].sort((a, b) => Number(Boolean(b.self)) - Number(Boolean(a.self)));
			this.render();
		}
		private get max(): number {
			const n = Number(this.getAttribute('max') ?? 4);
			return Number.isFinite(n) && n >= 1 ? Math.floor(n) : 4;
		}
		private render(): void {
			const doc = this.ownerDocument;
			this.setAttribute('role', 'group');
			this.setAttribute(
				'aria-label',
				this.getAttribute('label') ?? `${this.people.length} collaborators`,
			);
			const shown = this.people.slice(0, this.max);
			const items = shown.map((person) => {
				const li = doc.createElement('li');
				const button = doc.createElement('button');
				button.type = 'button';
				button.textContent = initialsOf(person.name);
				button.dataset.id = person.id;
				button.dataset.status = person.status ?? 'active';
				if (person.color) button.style.setProperty('--office-avatar-color', person.color);
				const status = person.status && person.status !== 'active' ? `, ${person.status}` : '';
				button.setAttribute('aria-label', `${person.name}${person.self ? ' (you)' : ''}${status}`);
				button.title = person.name;
				button.addEventListener('click', () =>
					emit(this, 'office-presence-select', { id: person.id }),
				);
				li.append(button);
				return li;
			});
			const hidden = this.people.length - shown.length;
			if (hidden > 0) {
				const li = doc.createElement('li');
				const chip = doc.createElement('span');
				chip.className = 'more';
				chip.textContent = `+${hidden}`;
				chip.title = this.people
					.slice(shown.length)
					.map((p) => p.name)
					.join(', ');
				chip.setAttribute('aria-label', `${hidden} more: ${chip.title}`);
				li.append(chip);
				items.push(li);
			}
			this.list.replaceChildren(...items);
		}
	}
	return OfficeUiPresence;
});
