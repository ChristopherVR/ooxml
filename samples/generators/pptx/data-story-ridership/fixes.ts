/**
 * Repairs to the saved package for two gaps of the core's writers, kept apart from the slides.
 *
 * Charts: a scatter series asked to draw no line (`line: 'none'`, `lineNoFill` on the model) is
 * still written with a solid `a:ln`, so PowerPoint joins the dots. The series outline is
 * rewritten to `a:noFill`.
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
const SERIES_LINE =
	/(<c:ser>(?:(?!<\/c:ser>).)*?<c:spPr>)<a:ln[^>]*><a:solidFill>(?:(?!<\/a:ln>).)*<\/a:solidFill><\/a:ln>(<\/c:spPr><c:marker>)/gsu;

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
			continue;
		}
		if (!/^ppt\/charts\/chart\d+\.xml$/u.test(name)) continue;
		const xml = decoder.decode(data);
		if (!xml.includes('<c:scatterChart>')) continue;
		const fixed = xml.replace(SERIES_LINE, '$1<a:ln w="19050"><a:noFill/></a:ln>$2');
		if (fixed === xml) throw new Error(`${name}: the scatter series outline was not found.`);
		changes[name] = fixed;
	}
	return patchPackage(bytes, changes);
}
