import JSZip from 'jszip';

/** Synthetic whole-shape numeric set with the native Transparency timing profile. */
export async function createOpacitySetDeck(hold: boolean): Promise<Buffer> {
	const { PresentationBuilder } = await import('../../core/builders/sdk/PresentationBuilder');
	const { handler, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addShape('rect', {
			x: 100,
			y: 100,
			width: 400,
			height: 200,
			fill: { type: 'solid', color: '#00AA00' },
		})
		.build();
	const element = slide.elements[0]!;
	if (element.type !== 'shape') throw new Error('Expected shape');
	element.text = 'Persistent text';
	element.textSegments = [{ text: element.text, style: { color: '#000000', fontSize: 24 } }];
	const next = createSlide('Blank')
		.addText('Next slide', { x: 100, y: 100, width: 400, height: 100 })
		.build();
	const bytes = await handler.save([slide, next]);
	handler.dispose();
	return patchOpacitySetPackage(bytes, hold);
}

/** Reuse the exact timing XML with source or published SDK fixture builders. */
export async function patchOpacitySetPackage(
	bytes: Uint8Array,
	hold: boolean,
	followingGroup = false,
): Promise<Buffer> {
	const zip = await JSZip.loadAsync(bytes);
	let xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const shapeId =
		/<p:cNvPr[^>]*\bid="(\d+)"[^>]*name="[^"\n]*Shape/u.exec(xml)?.[1] ??
		/<p:sp><p:nvSpPr><p:cNvPr[^>]*\bid="(\d+)"/u.exec(xml)?.[1];
	if (!shapeId) throw new Error('Missing serialized shape id');
	const fill = hold ? 'hold' : 'remove';
	const timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="6" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" grpId="0" presetID="9" presetClass="emph" presetSubtype="0" fill="${fill}" nodeType="clickEffect" dur="1000"><p:stCondLst><p:cond delay="${hold ? 1000 : 0}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="5" dur="1000" fill="${fill}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="0.4"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
	const nextGroup = `<p:par><p:cTn id="7" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="8" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="9" grpId="1" presetID="9" presetClass="emph" presetSubtype="0" fill="hold" nodeType="clickEffect" dur="1"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="10" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="0.4"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
	const completeTiming = followingGroup
		? timing.replace(
				'</p:childTnLst></p:cTn><p:prevCondLst>',
				nextGroup + '</p:childTnLst></p:cTn><p:prevCondLst>',
			)
		: timing;
	xml = xml.replace('</p:sld>', completeTiming + '</p:sld>');
	zip.file('ppt/slides/slide1.xml', xml);
	return zip.generateAsync({ type: 'nodebuffer' });
}
