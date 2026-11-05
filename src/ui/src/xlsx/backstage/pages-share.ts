/**
 * File > Share: whether this workbook is shared, who is in the session, and the "Share with people"
 * button that opens the Share dialog every product uses (`office-ui-share-dialog`), where the user
 * starts, joins or stops a session and copies its invitation link. Visio's page is the same.
 */
import { definePresence, type PresenceParticipant } from '../../presence';
import type { XlsxCollaborationState } from '../collaboration-types';
import { heading, paragraph, primary, type PageContext } from './parts';

/** What the Share page needs from the editor. */
export interface ShareHost {
	state(): XlsxCollaborationState;
	/** Share with people: the Share command (a host may take it over), else the Share dialog. */
	open(): void;
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

export function renderShare(page: PageContext): void {
	const { t, doc } = page;
	const share = page.host.share;
	if (!share) {
		page.content.replaceChildren(heading(page, t('Share')));
		return;
	}
	const state = share.state();
	const status = paragraph(
		page,
		t(
			state.status === 'connected' && !state.synced
				? 'Joining the session...'
				: STATUS[state.status],
		),
		'xve-share-status',
	);
	status.setAttribute('role', 'status');
	const children: Node[] = [
		heading(page, t('Share')),
		paragraph(
			page,
			t(
				'Edit this workbook together with other people. Everyone who joins the same session sees changes as they are made.',
			),
		),
	];
	if (state.active)
		children.push(paragraph(page, t('Session: {room}', { room: state.roomId ?? '' })));
	children.push(status);
	if (state.active) {
		definePresence();
		const people = doc.createElement('office-ui-presence') as PresenceElement;
		people.className = 'xve-share-people';
		people.setAttribute('label', t('People in this session'));
		people.participants = participants(state);
		children.push(people);
	}
	const open = primary(page, t('Share with people'), () => share.open());
	open.dataset.share = 'open';
	children.push(open);
	if (state.active) {
		const stop = primary(page, t('Stop sharing'), () => share.stop());
		stop.dataset.share = 'stop';
		children.push(stop);
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
