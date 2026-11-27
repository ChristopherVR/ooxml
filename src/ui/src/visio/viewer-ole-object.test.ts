import type { VisioOleObject, VisioShape } from 'ooxml-core/visio';
import { afterEach, expect, it, vi } from 'vitest';
import { oleObjectShape, openOleObject } from './viewer-ole-object';

vi.mock('ooxml-core/visio', async (original) => ({
	...(await original<typeof import('ooxml-core/visio')>()),
	readVisioEmbeddedObject: vi.fn(async () => ({
		bytes: new Uint8Array([0x50, 0x4b, 3, 4]),
		fileName: 'Microsoft_Word_Document.docx',
		format: 'docx',
	})),
}));

const word: VisioOleObject = {
	part: 'visio/embeddings/Microsoft_Word_Document.docx',
	format: 'docx',
	fileName: 'Microsoft_Word_Document.docx',
	showAsIcon: false,
};
const controller = { exportVsdx: () => ({ bytes: new Uint8Array([1]) }) } as never;

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

it('announces an embedded document as the cancelable openEmbeddedDocument file command', async () => {
	const host = document.body.appendChild(document.createElement('div'));
	const seen: unknown[] = [];
	host.addEventListener('file-command', (event) => {
		seen.push((event as CustomEvent).detail);
		event.preventDefault();
	});
	const click = vi.spyOn(HTMLAnchorElement.prototype, 'click');
	await openOleObject(host, controller, word, () => {});
	expect(seen).toEqual([
		{
			command: 'openEmbeddedDocument',
			fileName: 'Microsoft_Word_Document.docx',
			bytes: new Uint8Array([0x50, 0x4b, 3, 4]),
		},
	]);
	// The host took it, so nothing is downloaded.
	expect(click).not.toHaveBeenCalled();
});

it('downloads the embedded document when no host handles the command', async () => {
	const host = document.body.appendChild(document.createElement('div'));
	const click = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
	const create = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:test');
	await openOleObject(host, controller, word, () => {});
	expect(create).toHaveBeenCalledOnce();
	expect(click).toHaveBeenCalledOnce();
});

it('refuses an object it cannot open with the reason and raises no event', async () => {
	const host = document.body.appendChild(document.createElement('div'));
	const listener = vi.fn();
	host.addEventListener('file-command', listener);
	const messages: string[] = [];
	await openOleObject(
		host,
		controller,
		{ ...word, format: 'ole', refusal: 'Not here.' },
		(message) => messages.push(message),
	);
	expect(listener).not.toHaveBeenCalled();
	expect(messages).toEqual(['Not here.']);
});

it('finds an embedded object inside a group', () => {
	const inner = { id: '2', oleObject: word, children: [] } as unknown as VisioShape;
	const group = { id: '1', children: [inner] } as unknown as VisioShape;
	expect(oleObjectShape([group], '2')).toBe(inner);
	expect(oleObjectShape([group], '1')).toBeUndefined();
});
