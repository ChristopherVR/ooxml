import { describe, expect, it } from 'vitest';
import { createOneDriveClient } from './index';
import { decodeEntryId, encodeEntryId } from './browse';

const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
function setup(respond: (url: string) => Response | Promise<Response>) {
	const urls: string[] = [];
	const client = createOneDriveClient({
		getAccessToken: async () => 'tok',
		fetch: async (input) => {
			urls.push(String(input));
			return respond(String(input));
		},
	});
	return { client, urls };
}

describe('OneDrive browsing', () => {
	it('round-trips ids whose parts contain "!"', () => {
		const ref = { driveId: 'b!abc', itemId: 'ABC!123' };
		const id = encodeEntryId(ref);
		expect(id.split('!')).toHaveLength(2);
		expect(decodeEntryId(id)).toEqual(ref);
		expect(() => decodeEntryId('nonsense')).toThrow();
	});

	it('lists the root with a $select and maps entries', async () => {
		const { client, urls } = setup(() =>
			json({
				'@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=x',
				value: [
					{
						id: 'f1',
						name: 'Reports',
						folder: { childCount: 3 },
						parentReference: { driveId: 'd', path: '/drive/root:' },
					},
					{
						id: 'i1',
						name: 'Plan.docx',
						size: 12,
						file: {},
						lastModifiedDateTime: '2026-01-02T03:04:05Z',
						lastModifiedBy: { user: { displayName: 'Ada' } },
						parentReference: { driveId: 'd', path: '/drive/root:/Re%20ports/Q1' },
					},
					{
						id: 'n1',
						name: 'Notes',
						package: { type: 'oneNote' },
						parentReference: { driveId: 'd' },
					},
					{ name: 'no identity' },
				],
			}),
		);
		const listing = await client.list();
		expect(urls[0]).toMatch(/\/me\/drive\/root\/children\?\$select=.*&\$top=100$/u);
		expect(listing.next).toContain('skiptoken=x');
		expect(listing.entries).toHaveLength(3);
		expect(listing.entries[0]).toMatchObject({ type: 'folder', childCount: 3, name: 'Reports' });
		expect(listing.entries[1]).toMatchObject({
			type: 'file',
			kind: 'docx',
			size: 12,
			modified: Date.parse('2026-01-02T03:04:05Z'),
			modifiedBy: 'Ada',
			location: 'My files > Re ports > Q1',
		});
		expect(listing.entries[2]?.kind).toBeUndefined();
	});

	it('lists a folder by identity and follows a validated next link only', async () => {
		const { client, urls } = setup(() => json({ value: [] }));
		await client.list({ driveId: 'b!d', itemId: 'I!1' });
		expect(urls[0]).toContain('/drives/b!d/items/I!1/children?');
		await client.list(undefined, { page: 'https://graph.microsoft.com/v1.0/next?x=1' });
		expect(urls[1]).toBe('https://graph.microsoft.com/v1.0/next?x=1');
		for (const page of ['https://evil.example/x', 'http://graph.microsoft.com/x', 'x'])
			await expect(client.list(undefined, { page })).rejects.toMatchObject({
				code: 'invalid-response',
			});
		expect(urls).toHaveLength(2);
	});

	it('rejects a hostile next link from the server', async () => {
		const { client } = setup(() => json({ value: [], '@odata.nextLink': 'https://evil.example/' }));
		await expect(client.list()).rejects.toMatchObject({ code: 'invalid-response' });
	});

	it('resolves remoteItem identity for recent files', async () => {
		const { client, urls } = setup(() =>
			json({
				value: [
					{
						id: 'shortcut',
						name: 'Budget.xlsx',
						remoteItem: { id: 'real', file: {}, size: 5, parentReference: { driveId: 'other' } },
					},
				],
			}),
		);
		const { entries } = await client.recent();
		expect(urls[0]).toContain('/me/drive/recent?$select=');
		expect(decodeEntryId(entries[0]!.id)).toEqual({ driveId: 'other', itemId: 'real' });
		expect(entries[0]).toMatchObject({ kind: 'xlsx', size: 5 });
	});

	it('escapes single quotes and encodes search text; skips empty queries', async () => {
		const { client, urls } = setup(() => json({ value: [] }));
		await client.search("o'brien & 100% /x");
		expect(urls[0]).toContain("/me/drive/root/search(q='o''brien%20%26%20100%25%20%2Fx')?$select");
		await client.search('   ');
		expect(urls).toHaveLength(1);
	});
});
