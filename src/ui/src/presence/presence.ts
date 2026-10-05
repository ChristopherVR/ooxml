import { html } from 'lit';
import { OfficeElement, controlStyles } from '../base.js';
import { definer } from '../registry.js';
import css from './presence.css?raw';

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

/**
 * Presence avatar stack. `participants` property (array of `PresenceParticipant`), `max`
 * attribute (visible avatars, default 4; the rest collapse into a "+N" chip). `role="list"`
 * with one button per person (accessible name = name, status and "(you)"); activating one emits
 * `office-presence-select` `{ id }` so the host can follow that user. Rendering carries no
 * identity beyond what the host passes in.
 */
export class OfficeUiPresence extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		// `participants` keeps the local user first, so it has an accessor.
		max: { type: String },
		label: { type: String },
	};
	declare max: string | null;
	declare label: string | null;
	private people: PresenceParticipant[] = [];

	constructor() {
		super();
		this.max = null;
		this.label = null;
	}

	get participants(): PresenceParticipant[] {
		return this.people;
	}
	set participants(next: PresenceParticipant[]) {
		this.people = [...next].sort((a, b) => Number(Boolean(b.self)) - Number(Boolean(a.self)));
		this.requestUpdate('participants');
	}

	private get visible(): number {
		const n = Number(this.max ?? 4);
		return Number.isFinite(n) && n >= 1 ? Math.floor(n) : 4;
	}

	protected override willUpdate(): void {
		this.setAttribute('role', 'group');
		this.setAttribute('aria-label', this.label ?? `${this.people.length} collaborators`);
	}

	private person(person: PresenceParticipant) {
		const status = person.status && person.status !== 'active' ? `, ${person.status}` : '';
		return html`<li>
			<button
				type="button"
				data-id=${person.id}
				data-status=${person.status ?? 'active'}
				style=${person.color ? `--office-avatar-color:${person.color}` : ''}
				aria-label="${person.name}${person.self ? ' (you)' : ''}${status}"
				title=${person.name}
				@click=${() => this.fire('office-presence-select', { id: person.id })}
				>${initialsOf(person.name)}</button
			>
		</li>`;
	}

	protected override render() {
		const shown = this.people.slice(0, this.visible);
		const hidden = this.people.length - shown.length;
		const names = this.people
			.slice(shown.length)
			.map((p) => p.name)
			.join(', ');
		return html`<ul role="list">
			${shown.map((person) => this.person(person))}
			${
				hidden > 0
					? html`<li>
							<span class="more" title=${names} aria-label="${hidden} more: ${names}"
								>+${hidden}</span
							>
						</li>`
					: ''
			}
		</ul>`;
	}
}

export const definePresence = definer('office-ui-presence', () => OfficeUiPresence);
