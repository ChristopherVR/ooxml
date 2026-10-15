// The Calendar app's actions for the client: schedule, change and cancel meetings, and tie each to
// a conversation. A meeting with attendees gets a group chat (the meeting chat); a channel meeting
// posts a meeting card in the channel and uses the channel as its chat. Kept out of `store.ts`.
// DOM-free. New code.
import { type ChatMember, sanitizeMembers } from './chats';
import { meetingToIcs, icsFileName } from './ics';
import { type Meeting, type MeetingInput, type MeetingStore } from './meetings';
import type { TeamsWorkspace } from './workspace';

export interface MeetingActions {
	createMeeting: (input: MeetingFormInput) => Meeting | null;
	updateMeeting: (id: string, patch: Partial<MeetingFormInput>) => Meeting | null;
	cancelMeeting: (id: string) => boolean;
	/** The `.ics` file for a meeting, or null when it does not exist. */
	meetingIcs: (id: string) => { name: string; text: string } | null;
}

/** What the New meeting form submits: attendees are user ids, resolved through the roster. */
export interface MeetingFormInput extends Omit<MeetingInput, 'attendees'> {
	attendeeIds?: readonly string[];
}

export interface MeetingActionDeps {
	ws: TeamsWorkspace;
	meetings: MeetingStore;
	notice: (text: string) => void;
	refresh: () => void;
	/** Defaults to the clock; tests pass their own. */
	now?: () => number;
}

export const meetingChatId = (meetingId: string): string => `mt-${meetingId}`.slice(0, 64);

export function createMeetingActions(deps: MeetingActionDeps): MeetingActions {
	const { ws, meetings, notice, refresh } = deps;
	const attendees = (ids: readonly string[] = []): ChatMember[] => {
		const known = new Map(ws.roster.list().map((r) => [r.id, r]));
		return sanitizeMembers(
			ids.flatMap((id) => {
				const found = id === ws.user.id ? undefined : known.get(id);
				return found ? [{ id: found.id, name: found.name }] : [];
			}),
		);
	};
	const channelExists = (id: string): boolean =>
		ws.chat.channels().some((c) => c.id === id && c.kind === 'channel' && !c.archived);
	const card = (conversationId: string, meeting: Meeting): void => {
		ws.chat.post(conversationId, { text: `Meeting: ${meeting.title}`, meetingId: meeting.id });
	};
	const announce = (meeting: Meeting, text: string): void => {
		const target = meeting.chatId ?? meeting.channelId;
		if (target) ws.chat.post(target, { text, system: true });
	};
	const toInput = (form: Partial<MeetingFormInput>): Partial<MeetingInput> => {
		const { attendeeIds, ...rest } = form;
		return { ...rest, ...(attendeeIds ? { attendees: attendees(attendeeIds) } : {}) };
	};
	return {
		createMeeting(form) {
			const { attendeeIds, channelId: wanted, ...rest } = form;
			const channelId = wanted && channelExists(wanted) ? wanted : undefined;
			const meeting = meetings.create({
				...rest,
				attendees: channelId ? [] : attendees(attendeeIds),
				...(channelId ? { channelId } : {}),
			});
			if (!meeting) {
				notice('The meeting needs a title and an end after its start (24 hours at most)');
				return null;
			}
			let result = meeting;
			if (channelId) card(channelId, meeting);
			else if (meeting.attendees.length > 0) {
				const chat = ws.chat.createGroup({
					others: meeting.attendees,
					name: meeting.title,
					id: meetingChatId(meeting.id),
					meetingId: meeting.id,
				});
				if (chat) {
					result = meetings.update(meeting.id, { chatId: chat.id }) ?? meeting;
					card(chat.id, result);
				}
			}
			refresh();
			return result;
		},
		updateMeeting(id, form) {
			const current = meetings.get(id);
			if (!current) return null;
			const patch = toInput(form);
			delete patch.channelId;
			const updated = meetings.update(id, patch);
			if (!updated) {
				notice('Only the organizer can change a meeting');
				return null;
			}
			if (updated.chatId && patch.attendees) {
				const have = new Set(current.attendees.map((a) => a.id));
				const want = new Set(updated.attendees.map((a) => a.id));
				ws.chat.addPeople(
					updated.chatId,
					updated.attendees.filter((a) => !have.has(a.id)),
				);
				for (const a of current.attendees)
					if (!want.has(a.id)) ws.chat.removePerson(updated.chatId, a.id);
			}
			refresh();
			return updated;
		},
		cancelMeeting(id) {
			const meeting = meetings.get(id);
			const ok = meetings.cancel(id);
			if (ok && meeting) announce(meeting, `${ws.user.name} canceled "${meeting.title}"`);
			else notice('Only the organizer can cancel a meeting');
			refresh();
			return ok;
		},
		meetingIcs(id) {
			const meeting = meetings.get(id);
			return meeting
				? {
						name: icsFileName(meeting),
						text: meetingToIcs(meeting, { now: (deps.now ?? Date.now)() }),
					}
				: null;
		},
	};
}
