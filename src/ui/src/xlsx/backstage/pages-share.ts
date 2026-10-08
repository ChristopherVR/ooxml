/**
 * File > Share: start or stop co-editing this workbook in a named room and see who is in it. The
 * same panel as Visio's (room name, Start / Stop sharing, status, the shared presence stack); the
 * editor's `collaboration` API does the work.
 */
import { definePresence, type PresenceParticipant } from '../../presence';
import type { XlsxCollaborationState } from '../collaboration-types';
import { heading, labelled, paragraph, primary, type PageContext } from './parts';

/** What the Share page needs from the editor. */
export interface ShareHost {
	state(): XlsxCollaborationState;
	start(roomId: string): Promise<void>;
	stop(): void;
}

type PresenceElement = HTMLElement & { participants: PresenceParticipant[] };

const STATUS: Record<XlsxCollaborationState['status'], string> = {
	off: 'Not sharing.',
	connecting: 'Connecting...',
	connected: 'Connected.',
	disconnected: 'Disconnected. Changes are kept and sent when the connection returns.',
	error: 'The connection failed. Changes are kept and sent when the connection returns.',
};

/** A fresh room name a second window can type: `book-` and six letters or digits. */
export const suggestRoom = (): string => `book-${Math.random().toString(36).slice(2, 8)}`;

export function renderShare(page: PageContext): void {
	const { t, doc } = page;
	const share = page.host.share;
	if (!share) {
		page.content.replaceChildren(heading(page, t('Share')));
		return;
	}
	const state = share.state();
	const status = paragraph(page, t(STATUS[state.status]), 'xve-share-status');
	status.setAttribute('role', 'status');
	const say = (message: string) => {
		status.textContent = message;
	};
	const children: Node[] = [
		heading(page, t('Share')),
		paragraph(
			page,
			t(
				'Edit this workbook together with other people. Everyone who joins the same session sees changes as they are made.',
			),
		),
	];
	if (state.active) {
		children.push(paragraph(page, t('Session: {room}', { room: state.roomId ?? '' })));
		if (state.status === 'connected' && !state.synced) say(t('Joining the session...'));
	} else {
		const room = doc.createElement('input');
		room.type = 'text';
		room.className = 'xve-share-room';
		room.maxLength = 128;
		room.autocomplete = 'off';
		room.spellcheck = false;
		room.value = suggestRoom();
		children.push(labelled(page, t('Session name'), room));
		const start = primary(page, t('Start sharing'), () => {
			start.disabled = true;
			say(t('Connecting...'));
			share.start(room.value.trim()).catch((error: unknown) => {
				start.disabled = false;
				say(t(error instanceof Error ? error.message : String(error)));
			});
		});
		start.dataset.share = 'start';
		children.push(start);
	}
	children.push(status);
	if (state.active) {
		definePresence();
		const people = doc.createElement('office-ui-presence') as PresenceElement;
		people.className = 'xve-share-people';
		people.setAttribute('label', t('People in this session'));
		people.participants = participants(state);
		const stop = primary(page, t('Stop sharing'), () => share.stop());
		stop.dataset.share = 'stop';
		children.push(people, stop);
	}
	children.push(
		paragraph(
			page,
			t(
				'Cell values, formulas, formats, sheets, rows, columns, merges and names are shared. Charts, pictures, comments, conditional formats, data validation, tables and page setup stay in this window.',
			),
			'xve-backstage-muted',
		),
		paragraph(
			page,
			t(
				'Without a server set by the page, a session connects only windows and tabs of this browser. Nothing is uploaded.',
			),
			'xve-backstage-muted',
		),
	);
	page.content.replaceChildren(...children);
}

/** The people in the room as the shared presence stack lists them. */
export function participants(state: XlsxCollaborationState): PresenceParticipant[] {
	return state.people.map((person) => ({
		id: String(person.clientId),
		name: person.name,
		color: person.color,
		...(person.self ? { self: true } : {}),
	}));
}
