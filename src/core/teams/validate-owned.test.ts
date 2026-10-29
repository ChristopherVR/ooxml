import * as Y from 'yjs';
import { describe, expect, it } from 'vitest';
import { createChatStore } from './chat';
import { createTabStore } from './tabs';
import { validateTeamsUpdate, type UpdateActor } from './validate-update';

const ann: UpdateActor = { id: 'ann' };
const bob: UpdateActor = { id: 'bob' };
const mod: UpdateActor = { id: 'mod', moderator: true };

const copy = (from: Y.Doc): Y.Doc => {
	const doc = new Y.Doc();
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(from));
	return doc;
};

function capture(doc: Y.Doc, fn: () => void): Uint8Array {
	const updates: Uint8Array[] = [];
	const on = (u: Uint8Array): void => void updates.push(u);
	doc.on('update', on);
	fn();
	doc.off('update', on);
	return Y.mergeUpdates(updates);
}

const codes = (v: ReturnType<typeof validateTeamsUpdate>): string[] =>
	v.ok ? [] : v.violations.map((x) => x.code);

/** A server holding Ann's channel, tab and a pinned message. */
function seeded(): { server: Y.Doc; messageId: string; otherId: string; tabId: string } {
	const doc = new Y.Doc();
	// A low client id loses every concurrent write to the same key, which the seed test relies on.
	doc.clientID = 1;
	const store = createChatStore(doc, { id: 'ann', name: 'Ann' });
	store.createChannel({ id: 'general', name: 'General' });
	const message = store.post('general', { text: 'hello' })!;
	store.pin('general', message.id);
	const other = store.post('general', { text: 'second' })!;
	const tabs = createTabStore(doc, { id: 'ann', name: 'Ann' }, () => true);
	const tab = tabs.add('general', 'Docs', { type: 'website', url: 'https://example.com/' })!;
	return { server: doc, messageId: message.id, otherId: other.id, tabId: tab.id };
}

describe('validateTeamsUpdate: channels', () => {
	it('accepts a member creating a channel as themselves', () => {
		const { server } = seeded();
		const doc = copy(server);
		const store = createChatStore(doc, { id: 'bob', name: 'Bob' });
		const update = capture(doc, () => void store.createChannel({ name: 'Random' }));
		expect(validateTeamsUpdate(server, update, bob)).toEqual({ ok: true });
	});

	it('refuses a channel created in somebody else name', () => {
		const { server } = seeded();
		const doc = copy(server);
		const store = createChatStore(doc, { id: 'ann', name: 'Ann' });
		const update = capture(doc, () => void store.createChannel({ name: 'Forged' }));
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['forged-author']);
	});

	it('lets the creator archive and refuses everybody else', () => {
		const { server } = seeded();
		const doc = copy(server);
		const update = capture(
			doc,
			() => void createChatStore(doc, { id: 'x', name: 'X' }).archiveChannel('general'),
		);
		expect(validateTeamsUpdate(server, update, ann)).toEqual({ ok: true });
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['foreign-row']);
		expect(validateTeamsUpdate(server, update, mod)).toEqual({ ok: true });
	});

	it('refuses renaming or removing a channel of somebody else, allows a moderator', () => {
		const { server } = seeded();
		const rename = copy(server);
		const renamed = capture(rename, () =>
			rename.getMap<Y.Map<unknown>>('teams:channels').get('general')!.set('name', 'Mine'),
		);
		expect(codes(validateTeamsUpdate(server, renamed, bob))).toEqual(['foreign-row']);
		expect(validateTeamsUpdate(server, renamed, mod)).toEqual({ ok: true });
		const remove = copy(server);
		const removed = capture(remove, () => remove.getMap('teams:channels').delete('general'));
		expect(codes(validateTeamsUpdate(server, removed, bob))).toEqual(['foreign-row']);
		expect(validateTeamsUpdate(server, removed, ann)).toEqual({ ok: true });
	});

	it('does not let the creator hand a channel to somebody else', () => {
		const { server } = seeded();
		const doc = copy(server);
		const update = capture(doc, () =>
			doc.getMap<Y.Map<unknown>>('teams:channels').get('general')!.set('createdBy', 'bob'),
		);
		expect(codes(validateTeamsUpdate(server, update, ann))).toEqual(['author-changed']);
	});

	it('accepts a concurrent seed of the same default channel, not a renamed one', () => {
		const { server } = seeded();
		const same = new Y.Doc();
		const seed = capture(
			same,
			() =>
				void createChatStore(same, { id: 'bob', name: 'Bob' }).createChannel({
					id: 'general',
					name: 'General',
				}),
		);
		expect(validateTeamsUpdate(server, seed, bob)).toEqual({ ok: true });
		const other = new Y.Doc();
		other.clientID = 2 ** 31;
		const hijack = capture(
			other,
			() =>
				void createChatStore(other, { id: 'bob', name: 'Bob' }).createChannel({
					id: 'general',
					name: 'Mine',
				}),
		);
		expect(codes(validateTeamsUpdate(server, hijack, bob))).toEqual(['foreign-row']);
	});
});

describe('validateTeamsUpdate: tabs', () => {
	it('lets the creator rename and remove a tab, refuses others, allows a moderator', () => {
		const { server, tabId } = seeded();
		const doc = copy(server);
		const tabs = createTabStore(doc, { id: 'ann', name: 'Ann' }, () => true);
		const renamed = capture(doc, () => void tabs.rename(tabId, 'Renamed'));
		expect(validateTeamsUpdate(server, renamed, ann)).toEqual({ ok: true });
		expect(codes(validateTeamsUpdate(server, renamed, bob))).toEqual(['foreign-row']);
		expect(validateTeamsUpdate(server, renamed, mod)).toEqual({ ok: true });
		const again = copy(server);
		const removed = capture(
			again,
			() => void createTabStore(again, { id: 'ann', name: 'Ann' }, () => true).remove(tabId),
		);
		expect(codes(validateTeamsUpdate(server, removed, bob))).toEqual(['foreign-row']);
	});

	it('refuses a tab created in somebody else name, accepts one created as yourself', () => {
		const { server } = seeded();
		const doc = copy(server);
		const tabs = createTabStore(doc, { id: 'ann', name: 'Ann' }, () => true);
		const update = capture(
			doc,
			() => void tabs.add('general', 'New', { type: 'website', url: 'https://example.com/' }),
		);
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['forged-author']);
		expect(validateTeamsUpdate(server, update, ann)).toEqual({ ok: true });
	});
});

describe('validateTeamsUpdate: pins', () => {
	it('accepts pinning as yourself and refuses a pin in somebody else name', () => {
		const { server, otherId } = seeded();
		const doc = copy(server);
		const update = capture(
			doc,
			() => void createChatStore(doc, { id: 'ann', name: 'Ann' }).pin('general', otherId),
		);
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['forged-author']);
		expect(validateTeamsUpdate(server, update, ann)).toEqual({ ok: true });
	});

	it('lets any member unpin', () => {
		const { server, messageId } = seeded();
		const doc = copy(server);
		const update = capture(
			doc,
			() => void createChatStore(doc, { id: 'bob', name: 'Bob' }).unpin('general', messageId),
		);
		expect(validateTeamsUpdate(server, update, bob)).toEqual({ ok: true });
	});

	it('keeps non-members out of a private chat pins', () => {
		const { server, messageId } = seeded();
		server.getMap('teams:chat-members').set('ann', true);
		const doc = copy(server);
		const update = capture(
			doc,
			() => void createChatStore(doc, { id: 'bob', name: 'Bob' }).unpin('general', messageId),
		);
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['not-member']);
		expect(validateTeamsUpdate(server, update, mod)).toEqual({ ok: true });
	});
});
