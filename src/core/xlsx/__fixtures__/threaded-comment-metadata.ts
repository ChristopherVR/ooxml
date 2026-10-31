import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { NS, buildXml, children, first, parseXml, type XmlElement } from '../../xml/index';
export const bytes = new Uint8Array(
	readFileSync(new URL('./excel-threaded-comments.xlsx', import.meta.url)),
);

export const mentioned = '@Parity Author';
export const mentionId = '{11111111-1111-4111-8111-111111111111}';
const secondPerson = '{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}';
export const threadFiles = (zip: JSZip) => zip.file(/^xl\/threadedComments\/.*\.xml$/);
export const messages = (root: XmlElement) => children(root, 'threadedComment', NS.tc);
export const id = (node: XmlElement) => node.getAttribute('id');
export const sourceRoot = async (zip: JSZip) =>
	parseXml(await threadFiles(zip)[0]!.async('string')).documentElement;

function extension(node: XmlElement, value: string): void {
	const doc = node.ownerDocument!;
	const list = doc.createElementNS(NS.tc, 'extLst');
	const ext = doc.createElementNS(NS.x, 'x:ext');
	ext.setAttribute('uri', `urn:parity:${value}`);
	const kept = doc.createElementNS('urn:parity', 'p:kept');
	kept.textContent = value;
	ext.appendChild(kept);
	list.appendChild(ext);
	node.appendChild(list);
}

/** Native structure augmented with valid mention fields and opaque extension payloads. */
export async function metadataFixture() {
	const zip = await JSZip.loadAsync(bytes);
	const root = await sourceRoot(zip);
	const thread = messages(root);
	const comment = thread[0]!;
	comment.setAttribute('done', '1');
	first(comment, 'text', NS.tc)!.textContent = `Hello ${mentioned}!`;
	const doc = root.ownerDocument!;
	const mentions = doc.createElementNS(NS.tc, 'mentions');
	const mention = doc.createElementNS(NS.tc, 'mention');
	for (const [key, value] of Object.entries({
		mentionpersonId: comment.getAttribute('personId')!,
		mentionId,
		startIndex: '6',
		length: String(mentioned.length),
	}))
		mention.setAttribute(key, value);
	mentions.appendChild(mention);
	comment.appendChild(mentions);
	extension(comment, 'message');
	extension(root, 'container');
	thread[2]!.setAttribute('personId', secondPerson);
	zip.file(threadFiles(zip)[0]!.name, buildXml(doc));
	const persons = parseXml(await zip.file('xl/persons/person.xml')!.async('string'));
	const person = children(persons.documentElement, 'person', NS.tc)[0]!;
	person.setAttribute('userId', 'original-account');
	extension(person, 'person');
	const second = person.cloneNode(true) as XmlElement;
	second.setAttribute('id', secondPerson);
	second.setAttribute('userId', 'second-account');
	persons.documentElement.appendChild(second);
	zip.file('xl/persons/person.xml', buildXml(persons));
	return zip.generateAsync({ type: 'uint8array' });
}

