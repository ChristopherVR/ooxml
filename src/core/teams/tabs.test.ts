import * as Y from 'yjs';
import { createTabStore } from './tabs.js';

const website = { type: 'website' as const, url: 'https://example.com/' };
const docs: Y.Doc[] = [];
const doc = () => {
	const value = new Y.Doc();
	docs.push(value);
	return value;
};
const merge = (a: Y.Doc, b: Y.Doc) => {
	Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
	Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
};
afterEach(() => {
	for (const value of docs.splice(0)) value.destroy();
});

describe('shared channel tabs', () => {
	it('converges concurrent tabs and survives a document snapshot', () => {
		const a = doc();
		const b = doc();
		const ada = createTabStore(a, { id: 'ada', name: 'Ada' }, () => true);
		const bob = createTabStore(b, { id: 'bob', name: 'Bob' }, () => true);
		ada.add('general', 'Project site', website);
		bob.add('general', 'Budget', {
			type: 'file',
			attachment: { name: 'Budget.xlsx', kind: 'xlsx', url: '/files/budget.xlsx' },
		});
		merge(a, b);
		expect(ada.tabs('general')).toEqual(bob.tabs('general'));
		expect(ada.tabs('general')).toHaveLength(2);
		const restored = doc();
		Y.applyUpdate(restored, Y.encodeStateAsUpdate(a));
		expect(
			createTabStore(restored, { id: 'ada', name: 'Ada' }, () => true).tabs('general'),
		).toEqual(ada.tabs('general'));
	});
	it('checks owner actions and propagates rename and removal', () => {
		const a = doc();
		const b = doc();
		const ada = createTabStore(a, { id: 'ada@example.com', name: 'Ada' }, () => true);
		const bob = createTabStore(b, { id: 'bob', name: 'Bob' }, () => true);
		const tab = ada.add('general', 'Site', website)!;
		merge(a, b);
		expect(bob.rename(tab.id, 'Hijack')).toBe(false);
		expect(bob.remove(tab.id)).toBe(false);
		expect(ada.rename(tab.id, '<b>Project</b>')).toBe(true);
		merge(a, b);
		expect(bob.tabs('general')[0]?.name).toBe('Project');
		expect(ada.remove(tab.id)).toBe(true);
		merge(a, b);
		expect(bob.tabs('general')).toEqual([]);
	});
	it('rejects unknown channels, unsafe URLs, and name-only files', () => {
		const store = createTabStore(doc(), { id: 'a', name: 'Ada' }, (id) => id === 'general');
		expect(store.add('missing', 'Site', website)).toBeNull();
		expect(store.add('general', '', website)).toBeNull();
		expect(
			store.add('general', 'Unsafe', { type: 'website', url: 'javascript:alert(1)' }),
		).toBeNull();
		expect(
			store.add('general', 'No bytes', {
				type: 'file',
				attachment: { name: 'Budget.xlsx', kind: 'xlsx' },
			}),
		).toBeNull();
	});
	it('drops malformed peer data and hides tabs of archived channels', () => {
		const document = doc();
		let archived = false;
		const store = createTabStore(document, { id: 'a', name: 'Ada' }, () => !archived);
		store.add('general', 'Site', website);
		document.getMap('teams:tabs').set('bad', { name: 'Untrusted' });
		expect(store.tabs('general')).toHaveLength(1);
		archived = true;
		expect(store.tabs('general')).toEqual([]);
	});
});
