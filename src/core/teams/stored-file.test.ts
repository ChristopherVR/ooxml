import { storeAttachment } from './stored-file';
import type { UploadableFile } from './store';
it('preserves the original name for storage hosts while using unique remote keys', async () => {
	const file = Object.assign(new Blob(['pdf bytes'], { type: 'application/pdf' }), {
		name: 'brief.pdf',
	});
	const received: (UploadableFile & Blob)[] = [];
	const upload = async (item: UploadableFile & Blob) => {
		received.push(item);
		return { name: item.name, kind: 'other' as const, url: 'https://suite.test/#/attachment/test' };
	};
	const first = await storeAttachment(file, 'message', upload);
	await storeAttachment(file, 'message', upload);
	expect(first.name).toBe('brief.pdf');
	expect(received[0]!.originalName).toBe('brief.pdf');
	expect(received[0]!.name).not.toBe(received[1]!.name);
	expect(await received[0]!.text()).toBe('pdf bytes');
});
