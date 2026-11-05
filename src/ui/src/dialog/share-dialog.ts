import { html, nothing } from 'lit';
import {
	COLLABORATION_OFF_STATE,
	type CollaborationOptions,
	type CollaborationState,
	isValidRoomId,
	joinShareOptions,
	type ShareFields,
	type ShareLocation,
	shareInvitation,
	startShareOptions,
	suggestRoomId,
} from 'ooxml-core/collab';
import { OfficeElement, controlStyles, flag } from '../base';
import { initialsOf } from '../presence/presence';
import { definer } from '../registry';
import { defineDialog } from './dialog';
import css from './share-dialog.css?raw';
import { type OfficeShareDialogLabels, SHARE_DIALOG_LABELS } from './share-dialog-locales';

/** `office-share-start`: the user asked to start or join; the host sets `collaboration`. */
export type OfficeShareStartEvent = CustomEvent<{ options: CollaborationOptions }>;

type View = 'create' | 'join';

/**
 * `<office-ui-share-dialog>`: the Share dialog of every product. Not sharing, it offers two tabs:
 * start a session (session name, your name, an optional server) or join one (an invitation link or
 * a session name). Sharing, it shows the connection, the invitation link with Copy, the session's
 * server and the people in it, and Stop sharing. The form logic is `ooxml-core/collab`'s share
 * helpers; the session belongs to the host editor, which sets `state` from its
 * `collaborationState` after every `collaboration-change`.
 *
 * Properties: `open`, `state`, `defaults` (`{ roomId, userName, serverUrl }`; a missing room gets a
 * fresh `room-prefix` name), `invitation` (opens on the join tab with it filled in), `location`
 * (the page invitation links point at; `window.location` by default), `labels` (translated, see
 * `shareDialogLabels`), `error` (shown under the form, e.g. a failed start), `note` (what the
 * product shares) and `no-server` (hides the server field: the host decides it).
 * Events: `office-share-start` `{ options }`, `office-share-stop` and `office-share-close`.
 */
export class OfficeUiShareDialog extends OfficeElement {
	static startEvent = 'office-share-start';
	static stopEvent = 'office-share-stop';
	static closeEvent = 'office-share-close';
	static override styles = controlStyles(css);
	static override properties = {
		open: flag,
		state: { attribute: false },
		defaults: { attribute: false },
		invitation: { type: String },
		location: { attribute: false },
		labels: { attribute: false },
		error: { type: String },
		note: { type: String },
		noServer: { attribute: 'no-server', ...flag },
		roomPrefix: { attribute: 'room-prefix', type: String },
		view: { state: true },
		fields: { state: true },
		joinText: { state: true },
		copied: { state: true },
	};
	declare open: boolean;
	declare state: CollaborationState;
	declare defaults: Partial<ShareFields>;
	declare invitation: string;
	declare location: ShareLocation | null;
	declare labels: Partial<OfficeShareDialogLabels>;
	declare error: string;
	declare note: string;
	declare noServer: boolean;
	declare roomPrefix: string;
	declare view: View;
	declare fields: ShareFields;
	declare joinText: string;
	declare copied: boolean;
	/** The server of the session this dialog started, for its invitation link. */
	private sessionServer: string | null = null;
	private wasOpen = false;

	constructor() {
		super();
		this.open = false;
		this.state = COLLABORATION_OFF_STATE;
		this.defaults = {};
		this.invitation = '';
		this.location = null;
		this.labels = {};
		this.error = '';
		this.note = '';
		this.noServer = false;
		this.roomPrefix = 'room';
		this.view = 'create';
		this.fields = { roomId: '', userName: '', serverUrl: '' };
		this.joinText = '';
		this.copied = false;
	}

	private get text(): OfficeShareDialogLabels {
		return { ...SHARE_DIALOG_LABELS.en, ...this.labels };
	}

	/** Seed the form each time the dialog opens, so it never shows a stale session name. */
	protected override willUpdate(): void {
		if (this.open && !this.wasOpen) {
			const defaults = this.defaults ?? {};
			this.fields = {
				roomId: defaults.roomId || suggestRoomId(this.roomPrefix || 'room'),
				userName: defaults.userName ?? '',
				serverUrl: defaults.serverUrl ?? '',
			};
			this.joinText = this.invitation;
			this.view = this.invitation ? 'join' : 'create';
			this.copied = false;
		}
		this.wasOpen = this.open;
		if (!this.state?.active) this.sessionServer = null;
	}

	/** The options the current form describes, or null while it is incomplete. */
	formOptions(): CollaborationOptions | null {
		if (this.view === 'join') {
			return joinShareOptions({
				invitation: this.joinText,
				userName: this.fields.userName,
				serverUrl: this.fields.serverUrl,
			});
		}
		return startShareOptions(this.fields);
	}

	/** The invitation link of the live session, or '' when not sharing. */
	invitationLink(): string {
		const roomId = this.state?.roomId;
		if (!this.state?.active || !roomId) return '';
		const serverUrl = this.sessionServer ?? this.defaults?.serverUrl ?? '';
		const page = this.location ?? this.ownerDocument.defaultView?.location;
		return shareInvitation(
			{ roomId, serverUrl },
			page ? { origin: page.origin, pathname: page.pathname } : undefined,
		);
	}

	private submit(event: Event): void {
		event.preventDefault();
		const options = this.formOptions();
		if (!options) return;
		this.sessionServer = options.serverUrl ?? '';
		this.fire((this.constructor as typeof OfficeUiShareDialog).startEvent, { options });
	}

	private close(): void {
		this.open = false;
		this.fire((this.constructor as typeof OfficeUiShareDialog).closeEvent, {});
	}

	private async copy(): Promise<void> {
		const link = this.invitationLink();
		try {
			await this.ownerDocument.defaultView?.navigator.clipboard.writeText(link);
			this.copied = true;
		} catch {
			// Without clipboard access, select the link so the user can copy it themselves.
			this.renderRoot.querySelector<HTMLInputElement>('.link input')?.select();
		}
	}

	private edit(key: keyof ShareFields, event: Event): void {
		this.fields = { ...this.fields, [key]: (event.target as HTMLInputElement).value };
	}

	private status(): string {
		const t = this.text;
		const state = this.state ?? COLLABORATION_OFF_STATE;
		if (state.status === 'connected') return state.synced ? t.statusConnected : t.statusJoining;
		if (state.status === 'connecting') return t.statusConnecting;
		if (state.status === 'disconnected') return t.statusDisconnected;
		if (state.status === 'error') return t.statusError;
		return t.statusOff;
	}

	private field(label: string, key: keyof ShareFields, placeholder = '') {
		return html`<label class="field"
			><span>${label}</span
			><input
				type="text"
				data-key=${key}
				autocomplete="off"
				spellcheck="false"
				maxlength=${key === 'roomId' ? 128 : 256}
				placeholder=${placeholder}
				.value=${this.fields[key]}
				@input=${(event: Event) => this.edit(key, event)}
		/></label>`;
	}

	private renderForm() {
		const t = this.text;
		const room = this.fields.roomId.trim();
		const badRoom = this.view === 'create' && room.length > 0 && !isValidRoomId(room);
		const ready = this.formOptions() !== null;
		const tab = (view: View, label: string) =>
			html`<button
				type="button"
				role="tab"
				class="tab"
				data-view=${view}
				aria-selected=${String(this.view === view)}
				@click=${() => (this.view = view)}
				>${label}</button
			>`;
		return html`<form class="body" @submit=${(event: Event) => this.submit(event)}>
			<p class="description">${t.description}</p>
			<div class="tabs" role="tablist">${tab('create', t.createTab)}${tab('join', t.joinTab)}</div>
			${
				this.view === 'join'
					? html`<label class="field"
							><span>${t.invitation}</span
							><input
								type="text"
								data-key="invitation"
								autocomplete="off"
								spellcheck="false"
								placeholder=${t.invitationPlaceholder}
								.value=${this.joinText}
								@input=${(event: Event) => (this.joinText = (event.target as HTMLInputElement).value)}
						/></label>`
					: this.field(t.roomName, 'roomId')
			}${badRoom ? html`<p class="error" role="alert">${t.invalidRoom}</p>` : nothing}
			${this.field(t.yourName, 'userName')}
			${
				this.noServer
					? nothing
					: html`${this.field(t.serverUrl, 'serverUrl', t.serverPlaceholder)}<p class="hint"
								>${t.serverHint}</p
							>`
			}
			${this.error ? html`<p class="error" role="alert">${this.error}</p>` : nothing}
			${this.note ? html`<p class="hint">${this.note}</p>` : nothing}
			<div class="actions">
				<button type="button" class="cancel" @click=${() => this.close()}>${t.cancel}</button
				><button type="submit" class="primary" data-share="start" ?disabled=${!ready}
					>${this.view === 'join' ? t.join : t.start}</button
				>
			</div>
		</form>`;
	}

	private renderActive() {
		const t = this.text;
		const state = this.state;
		const people = state.people ?? [];
		const server = this.sessionServer ?? this.defaults?.serverUrl ?? '';
		return html`<div class="body">
			<p class="status" role="status" data-status=${state.status}
				><span class="dot" aria-hidden="true"></span>${this.status()}<span class="count"
					>${t.peopleCount.replace('{count}', String(people.length))}</span
				></p
			>
			<div class="link field">
				<span>${t.link}</span>
				<div class="copy-row">
					<input type="text" readonly aria-label=${t.link} .value=${this.invitationLink()} /><button
						type="button"
						class="copy"
						@click=${() => void this.copy()}
						>${this.copied ? t.copied : t.copy}</button
					>
				</div>
				<p class="hint">${t.linkHint}</p>
			</div>
			<dl class="details">
				<dt>${t.session}</dt>
				<dd><code>${state.roomId ?? ''}</code></dd>
				<dt>${t.server}</dt>
				<dd>${server ? html`<code>${server}</code>` : t.noServer}</dd>
			</dl>
			<div class="people">
				<span class="people-label">${t.people}</span>
				<ul aria-label=${t.people}>
					${people.map(
						(person) =>
							html`<li
								><span
									class="avatar"
									aria-hidden="true"
									style=${`--share-person:${person.color || 'var(--office-accent)'}`}
									>${initialsOf(person.name)}</span
								><span class="name">${person.name}</span
								>${person.self ? html`<span class="you">${t.you}</span>` : nothing}</li
							>`,
					)}
				</ul>
			</div>
			${this.error ? html`<p class="error" role="alert">${this.error}</p>` : nothing}
			${this.note ? html`<p class="hint">${this.note}</p>` : nothing}
			<div class="actions">
				<button
					type="button"
					class="danger"
					data-share="stop"
					@click=${() => this.fire((this.constructor as typeof OfficeUiShareDialog).stopEvent, {})}
					>${t.stop}</button
				><button type="button" class="primary" @click=${() => this.close()}>${t.close}</button>
			</div>
		</div>`;
	}

	protected override render() {
		defineDialog(this.ownerDocument.defaultView?.customElements);
		const t = this.text;
		const active = !!this.state?.active;
		return html`<office-ui-dialog
			?open=${this.open}
			heading=${active ? t.activeTitle : t.title}
			close-label=${t.close}
			@office-dialog-close=${(event: Event) => {
				event.stopPropagation();
				this.close();
			}}
			>${active ? this.renderActive() : this.renderForm()}</office-ui-dialog
		>`;
	}
}

export const defineShareDialog = definer('office-ui-share-dialog', () => OfficeUiShareDialog);
