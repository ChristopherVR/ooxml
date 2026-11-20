/**
 * The security-bar contract every binding's building blocks pass: a deck that links a picture to
 * the web opens with the link in its model but nothing drawn from it, the blocks report
 * `externalContent.barVisible`, the binding's own `ExternalContentBar` (mounted by the harness
 * from that state) is the only bar on the page, and its Enable Content draws the linked picture
 * without changing the deck. A live host `allowExternalContent: false` hides the bar. Each
 * binding calls {@link runExternalContentContract} from one test with the harness it already has
 * for `runBuildingBlocksContract`.
 */
import { buildingBlocksContractDeck } from './building-blocks-contract.mjs';

export const LINKED_CONTENT_ORIGIN = 'https://linked-content.test';

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const LINKED_FILL = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;

/** The contract deck with a shape on its first slide filled from a picture at the linked origin. */
export async function linkedContentDeck(Handler, JSZip) {
	const zip = await JSZip.loadAsync(await buildingBlocksContractDeck(Handler));
	const slidePath = 'ppt/slides/slide1.xml';
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const slide = await zip.file(slidePath).async('string');
	zip.file(slidePath, slide.replace('</p:spTree>', `${LINKED_FILL}</p:spTree>`));
	const rels = await zip.file(relsPath).async('string');
	zip.file(
		relsPath,
		rels.replace(
			'</Relationships>',
			`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${LINKED_CONTENT_ORIGIN}/fill.png" TargetMode="External"/></Relationships>`,
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function until(harness, check, label) {
	const deadline = Date.now() + 15000;
	for (;;) {
		await harness.settle();
		let failure;
		try {
			check();
			return;
		} catch (error) {
			failure = error;
		}
		if (Date.now() > deadline) {
			throw new Error(`${harness.name}: timed out waiting for ${label}: ${String(failure)}`);
		}
		await new Promise((resolve) => setTimeout(resolve, 20));
	}
}

export async function runExternalContentContract({ harness, expect, Handler, JSZip }) {
	const host = document.createElement('div');
	document.body.append(host);
	const content = await linkedContentDeck(Handler, JSZip);
	const options = (allowExternalContent) => ({
		content,
		canEdit: false,
		onRequest: () => undefined,
		onDirtyChange: () => undefined,
		...(allowExternalContent === undefined ? {} : { allowExternalContent }),
	});
	const bars = () => host.querySelectorAll('pptx-ui-external-content-bar');
	const modelLinks = () => JSON.stringify(harness.slides()).includes(LINKED_CONTENT_ORIGIN);
	/** Whether the mounted blocks draw anything that would contact the linked server. */
	const drawn = () => host.innerHTML.includes(LINKED_CONTENT_ORIGIN);
	try {
		await harness.mount(host, options(undefined));
		await until(harness, () => expect(harness.state().slideCount).toBe(2), 'the deck');
		await until(harness, () => expect(harness.externalContent().barVisible).toBe(true), 'the bar');
		// The host shell's bar is the only one: the canvas leaves it to the shell, as in React.
		await until(harness, () => expect(bars()).toHaveLength(1), 'one bar element');
		expect(modelLinks()).toBe(true);
		expect(drawn()).toBe(false);
		expect(harness.externalContent().allowed).toBe(false);
		const slides = JSON.stringify(harness.slides());

		await harness.update(options(false));
		await until(
			harness,
			() => expect(harness.externalContent().barVisible).toBe(false),
			'host false',
		);
		await until(harness, () => expect(bars()).toHaveLength(0), 'no bar for host false');
		await harness.update(options(undefined));
		await until(harness, () => expect(bars()).toHaveLength(1), 'the bar again');

		bars()[0].dispatchEvent(
			new CustomEvent('external-content-request', { detail: { id: 'action' }, bubbles: true }),
		);
		await until(harness, () => expect(drawn()).toBe(true), 'the linked picture');
		expect(harness.externalContent().barVisible).toBe(false);
		expect(harness.externalContent().allowed).toBe(true);
		expect(bars()).toHaveLength(0);
		expect(harness.state().slideCount).toBe(2);
		// Enable Content changed only what is drawn: the deck is the one that was open.
		expect(JSON.stringify(harness.slides())).toBe(slides);
	} finally {
		harness.destroy();
		host.remove();
	}
}
