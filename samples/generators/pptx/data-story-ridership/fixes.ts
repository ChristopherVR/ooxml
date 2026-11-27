/**
 * Repairs to the saved package for a gap of the core's writers, kept apart from the slides.
 *
 * SmartArt: the drawing the core lays out for the chevron process sets the node text at a size
 * that does not fit the chevrons and leaves it left-aligned, and a node's own run style is not
 * applied. The cached drawing's text is set to a size that fits, bold, centred and inset so it
 * stays on the body of each chevron.
 */
import { patchPackage, readPackage } from '../../lib/package';

const decoder = new TextDecoder();
/** The depth of a chevron's notch and point at this size, in EMU: text stays between them. */
const INSET = 571500;

export async function repairPackage(bytes: Uint8Array): Promise<Uint8Array> {
	const changes: Record<string, string> = {};
	for (const [name, data] of await readPackage(bytes)) {
		if (/^ppt\/diagrams\/drawing\d+\.xml$/u.test(name)) {
			changes[name] = decoder
				.decode(data)
				.replaceAll(
					'<a:bodyPr anchor="ctr"/>',
					`<a:bodyPr lIns="${INSET}" rIns="${INSET}" anchor="ctr"/>`,
				)
				.replaceAll('<a:pPr/>', '<a:pPr algn="ctr"/>')
				.replace(/(<a:rPr lang="en-US") sz="\d+"/gu, '$1 sz="1400" b="1"');
		}
	}
	return patchPackage(bytes, changes);
}
