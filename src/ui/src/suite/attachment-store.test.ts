import 'fake-indexeddb/auto';
import { SuiteAttachmentStore } from './attachment-store';
it('stores general file bytes, preserves names and isolates local profiles', async () => {
	const name = `attachments-${crypto.randomUUID()}`;
	const store = new SuiteAttachmentStore(name);
	for (const filename of ['brief.pdf', 'image.png', 'archive.zip', 'no-extension']) {
		const bytes = new Uint8Array([0, 255, 17, 128]);
		const file = await store.create(filename, 'application/octet-stream', bytes);
		const restored = await new SuiteAttachmentStore(name).get(file.id);
		expect(restored.name).toBe(filename);
		expect([...restored.bytes]).toEqual([...bytes]);
		await expect(new SuiteAttachmentStore(name + '-other').get(file.id)).rejects.toThrow(
			'current local profile',
		);
	}
});
