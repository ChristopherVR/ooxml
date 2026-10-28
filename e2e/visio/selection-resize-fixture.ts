import { expect, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';

export async function fixture(locked = false): Promise<Buffer> {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Selection resize'));
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	const template = xml.match(/<Shape ID="1"[\s\S]*?<\/Shape>/u)![0];
	const shapes = ['1', '2']
		.map((id) =>
			template
				.replace('ID="1"', `ID="${id}" Type="Shape"`)
				.replace('N="PinX" V="4"', `N="PinX" V="${id === '1' ? 2 : 5}"`)
				.replace('N="Width" V="3"', 'N="Width" V="1"')
				.replaceAll('<Cell N="X" V="3"/>', '<Cell N="X" V="1" F="Width"/>')
				.replaceAll('<Cell N="Y" V="1"/>', '<Cell N="Y" V="1" F="Height"/>')
				.replace(
					'<Text>',
					'<Cell N="LocPinX" V="0.5" F="Width*0.5"/><Cell N="LocPinY" V="0.5" F="Height*0.5"/><Cell N="Angle" V="0"/><Text>',
				)
				.replace('<Text>', `${locked && id === '2' ? '<Cell N="LockWidth" V="1"/>' : ''}<Text>`),
		)
		.join('');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(/<Shapes>[\s\S]*?<\/Shapes>/u, `<Shapes>${shapes}</Shapes>`),
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

export const inventory = async (viewer: Locator) => {
	await expect
		.poll(() =>
			viewer.evaluate((node) => {
				const state = (node as VisioViewerElement).controller.state;
				return state.loading || state.edit.busy;
			}),
		)
		.toBe(false);
	return viewer.evaluate((node) => {
		const controller = (node as VisioViewerElement).controller;
		return {
			busy: controller.state.loading || controller.state.edit.busy,
			bytes: Array.from(controller.exportVsdx().bytes),
			shapes: controller.state.document!.pages[0]!.shapes.map((shape) => ({
				width: shape.width,
				height: shape.height,
				transform: [...shape.transform],
				pin: shape.rotation,
				children: shape.children.map((member) => ({
					id: member.id,
					width: member.width,
					height: member.height,
					transform: [...member.transform],
					pin: member.rotation,
				})),
			})),
			selection: controller.state.selectedShapes.map((shape) => shape.id),
			error: controller.state.edit.error?.message ?? '',
		};
	});
};
export const handlePoint = (viewer: Locator) =>
	viewer.locator('[data-resize-handle="ne"]').evaluate((node) => {
		const circle = node as SVGCircleElement;
		const point = new DOMPoint(circle.cx.baseVal.value, circle.cy.baseVal.value).matrixTransform(
			circle.getScreenCTM()!,
		);
		const ctm = circle.ownerSVGElement!.getScreenCTM()!;
		return { x: point.x, y: point.y, dx: ctm.a, dy: -0.5 * ctm.d };
	});
