import { expect, it } from 'vitest';
import { ViewerController } from './controller';
import { mountDiagramCanvas, mountDiagramInspector } from './building-blocks';
import { THEME_STYLE_ID } from '../theme';
it('standalone diagram surfaces install shared theme in their owner document', () => {
	const owner = document.implementation.createHTMLDocument('Embedded diagram');
	const canvas = owner.createElement('div');
	const inspector = owner.createElement('aside');
	owner.body.append(canvas, inspector);
	const controller = new ViewerController();
	const stopCanvas = mountDiagramCanvas(canvas, controller);
	const stopInspector = mountDiagramInspector(inspector, controller);
	try {
		expect(owner.head.querySelectorAll(`#${THEME_STYLE_ID}`)).toHaveLength(1);
		expect(inspector.dataset.officeBlock).toBe('inspector');
		expect(inspector.dataset.officeProduct).toBe('visio');
		expect(canvas.shadowRoot!.querySelector('[part="canvas"]')).not.toBeNull();
	} finally {
		stopInspector();
		stopCanvas();
		controller.destroy();
	}
	expect(canvas.shadowRoot!.children).toHaveLength(0);
	expect(inspector.children).toHaveLength(0);
});
