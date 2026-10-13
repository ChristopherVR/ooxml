import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';

/** Isolate a genuine native preset-zero entrance with an independent concurrent clock marker. */
export async function createVisibilityTimingDeck(delay: number): Promise<Buffer> {
	const source = await readFile(
		new URL('../fixtures/e2e/animation-after-effect.pptx', import.meta.url),
	);
	const zip = await JSZip.loadAsync(source);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const document = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false }).parse(
		xml,
	);
	const slide = document['p:sld'];
	const main = slide['p:timing']['p:tnLst']['p:par']['p:cTn']['p:childTnLst']['p:seq']['p:cTn'];
	const group = main['p:childTnLst']['p:par'][0];
	main['p:childTnLst']['p:par'] = group;
	const effect = group['p:cTn']['p:childTnLst']['p:par']['p:cTn']['p:childTnLst']['p:par']['p:cTn'];
	if (effect['@_presetClass'] !== 'entr' || effect['@_presetID'] !== '0')
		throw new Error('Unexpected genuine source effect');
	delete effect['p:subTnLst'];
	const behavior = effect['p:childTnLst']['p:set'];
	if (behavior['p:to']['p:strVal']['@_val'] !== 'visible')
		throw new Error('Unexpected visibility set');
	behavior['p:cBhvr']['p:cTn']['p:stCondLst']['p:cond']['@_delay'] = String(delay);
	const marker = structuredClone(effect);
	marker['@_id'] = '7';
	marker['@_nodeType'] = 'withEffect';
	const markerBehavior = marker['p:childTnLst']['p:set'];
	markerBehavior['p:cBhvr']['p:cTn']['@_id'] = '8';
	markerBehavior['p:cBhvr']['p:cTn']['p:stCondLst']['p:cond']['@_delay'] = '0';
	markerBehavior['p:cBhvr']['p:tgtEl']['p:spTgt']['@_spid'] = '4';
	group['p:cTn']['p:childTnLst']['p:par']['p:cTn']['p:childTnLst']['p:par'] = [
		{ 'p:cTn': effect },
		{ 'p:cTn': marker },
	];
	// Keep physical shapes and build-list references intact while isolating the effect.
	zip.file('ppt/slides/slide1.xml', new XMLBuilder({ ignoreAttributes: false }).build(document));
	return zip.generateAsync({ type: 'nodebuffer' });
}
