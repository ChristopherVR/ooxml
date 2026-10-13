import JSZip from 'jszip';
import { PresentationBuilder } from 'pptx-viewer-core';

/** Native PresentationML shape-paint target, leaving its text outside the background build. */
export async function backgroundAnimationDeck(
	emphasis = false,
	wholeOpacity = false,
): Promise<Buffer> {
	const { handler, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank')
		.addShape('rect', {
			x: 150,
			y: 140,
			width: 500,
			height: 220,
			fill: { type: 'solid', color: '#00AA00' },
		})
		.build();
	const element = slide.elements[0]!;
	if (element.type !== 'shape') throw new Error('Expected shape');
	element.text = 'Persistent text';
	element.textSegments = [{ text: element.text, style: { fontSize: 36, color: '#000000' } }];
	if (emphasis) {
		element.shapeStyle = { ...element.shapeStyle, strokeColor: '#0000FF', strokeWidth: 6 };
	}
	slide.animations = [
		{ elementId: element.id, entrance: 'fadeIn', durationMs: 1000, trigger: 'onClick' },
	];
	const zip = await JSZip.loadAsync(await handler.save([slide]));
	handler.dispose();
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	let patched = xml.replace(
		/<p:spTgt([^>]*?)(?:\/>|>\s*<\/p:spTgt>)/gu,
		'<p:spTgt$1><p:bg/></p:spTgt>',
	);
	if (patched === xml) throw new Error('No native animation shape target');
	if (emphasis) {
		const shapeId = /<p:spTgt spid="(\d+)"/u.exec(patched)?.[1];
		if (!shapeId) throw new Error('Missing serialized numeric shape id');
		const color = (id: number, attribute: string, to: string) =>
			`<p:animClr clrSpc="rgb" dir="cw"><p:cBhvr><p:cTn id="${id}" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="${shapeId}"><p:bg/></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>${attribute}</p:attrName></p:attrNameLst></p:cBhvr><p:to><a:srgbClr val="${to}"/></p:to></p:animClr>`;
		patched = patched
			.replace('presetClass="entr"', 'presetClass="emph"')
			.replace('presetID="10"', 'presetID="1"')
			.replace(/<p:set>[\s\S]*?<\/p:set>/u, '')
			.replace(
				/<p:animEffect[\s\S]*?<\/p:animEffect>/u,
				color(3, 'fillcolor', 'FF0000') + color(8, 'stroke.color', 'FF00FF'),
			);
		if (wholeOpacity) {
			patched = patched.replace(
				'</p:animClr></p:childTnLst></p:cTn></p:par>',
				'</p:animClr></p:childTnLst></p:cTn></p:par>' +
					`<p:par><p:cTn id="9" presetID="9" presetClass="emph" presetSubtype="0" fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="10" dur="1000" fill="hold"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="1"/></p:val></p:tav><p:tav tm="100000"><p:val><p:fltVal val="0.4"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par>`,
			);
		}
	}
	zip.file('ppt/slides/slide1.xml', patched);
	return zip.generateAsync({ type: 'nodebuffer' });
}
