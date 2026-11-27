/**
 * Fixes applied to the package after the core has saved it, for gaps in the core's writers that
 * this deck runs into. Each one is reported with the sample.
 */
import { FONTS } from '../../lib/fiction';
import { patchPackage, readPackage } from '../../lib/package';

const decoder = new TextDecoder();

/**
 * Removes the legend of the chart that has a series called `seriesName`. The core refuses to
 * write an embedded workbook for a chart without a legend, so the chart is saved with one and
 * the legend is taken out here.
 */
export async function removeChartLegend(
	bytes: Uint8Array,
	seriesName: string,
): Promise<Uint8Array> {
	const changes: Record<string, string> = {};
	for (const [name, data] of await readPackage(bytes)) {
		if (!/^ppt\/charts\/chart\d+\.xml$/u.test(name)) continue;
		const xml = decoder.decode(data);
		if (xml.includes(`>${seriesName}<`))
			changes[name] = xml.replace(/<c:legend>.*?<\/c:legend>/su, '');
	}
	if (Object.keys(changes).length !== 1)
		throw new Error(`Expected one chart with a series named "${seriesName}".`);
	return patchPackage(bytes, changes);
}

export interface DiagramStyle {
	/** Fill and text colour of each node, in order (`#RRGGBB`). */
	fills: readonly string[];
	text: readonly string[];
	/** Node text size in points. */
	size: number;
}

const colorList = (colors: readonly string[]) =>
	colors.map((color) => `<a:srgbClr val="${color.slice(1).toUpperCase()}"/>`).join('');

/**
 * Gives the deck's one SmartArt graphic its own colours and type. The core writes a node's
 * direct formatting (`style.fillColor`, `fontColor`, `bold`) only into the cached drawing part;
 * PowerPoint lays the diagram out again from the data, colour and style parts and so shows the
 * stock colour scheme. Here the colour part gets one fill and one text colour per node, and the
 * node text in the data part gets its weight, size and font, so both readings agree.
 */
export async function restyleDiagram(bytes: Uint8Array, style: DiagramStyle): Promise<Uint8Array> {
	const parts = await readPackage(bytes);
	const part = (name: string) => {
		const data = parts.get(name);
		if (!data) throw new Error(`The package has no ${name}.`);
		return decoder.decode(data);
	};
	const run = `<a:rPr lang="en-US" sz="${style.size * 100}" b="1"><a:latin typeface="${FONTS.display}"/></a:rPr>`;
	let node = 0;
	return patchPackage(bytes, {
		'ppt/diagrams/colors1.xml': part('ppt/diagrams/colors1.xml')
			.replace(
				/<dgm:fillClrLst meth="cycle">.*?<\/dgm:fillClrLst>/gu,
				`<dgm:fillClrLst meth="repeat">${colorList(style.fills)}</dgm:fillClrLst>`,
			)
			.replace(
				/<dgm:txFillClrLst meth="repeat">.*?<\/dgm:txFillClrLst>/gu,
				`<dgm:txFillClrLst meth="repeat">${colorList(style.text)}</dgm:txFillClrLst>`,
			),
		'ppt/diagrams/data1.xml': part('ppt/diagrams/data1.xml').replaceAll(
			'<a:rPr lang="en-US"/><a:t>',
			`${run}<a:t>`,
		),
		'ppt/diagrams/drawing1.xml': part('ppt/diagrams/drawing1.xml').replace(
			/<a:rPr lang="en-US" sz="\d+"[^>]*\/>/gu,
			() =>
				`<a:rPr lang="en-US" sz="${style.size * 100}" b="1"><a:solidFill>${colorList([style.text[node++ % style.text.length]!])}</a:solidFill><a:latin typeface="${FONTS.display}"/></a:rPr>`,
		),
	});
}
