// @vitest-environment happy-dom
import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, it, expect } from 'vitest';

import { ConnectorElementRenderer } from './ConnectorElementRenderer';
import type { ConnectorRendererProps } from './element-renderer-types';

/** PowerPoint's dash lengths and cap handling reach the React connector. */
function render(shapeStyle: Record<string, unknown>): string {
	const el = {
		id: 'c',
		type: 'connector',
		shapeType: 'straightConnector1',
		x: 0,
		y: 0,
		width: 100,
		height: 0,
		shapeStyle: { strokeColor: '#000000', strokeWidth: 6, ...shapeStyle },
	} as unknown as PptxElement;
	const props = {
		el,
		isSelected: false,
		canInteract: false,
		showResizeHandles: false,
		showHoverBorder: false,
		selectionColorClass: 'blue-500',
		opacity: 1,
		zIndex: 3,
	} as unknown as ConnectorRendererProps;
	return renderToStaticMarkup(<ConnectorElementRenderer {...props} />);
}

describe('ConnectorElementRenderer dash patterns', () => {
	it('uses the measured dash and gap lengths', () => {
		expect(render({ strokeDash: 'dash' })).toContain('stroke-dasharray="24 18"');
	});

	it('folds a round cap into the dash and keeps square dashes flat', () => {
		const round = render({ strokeDash: 'dash', lineCap: 'rnd' });
		expect(round).toContain('stroke-dasharray="18 24"');
		expect(round).toContain('stroke-linecap="round"');
		const square = render({ strokeDash: 'dash', lineCap: 'sq' });
		expect(square).toContain('stroke-dasharray="24 18"');
		expect(square).not.toContain('stroke-linecap="square"');
	});

	it("draws a dashed square-cap line with PowerPoint's overshoot at its start", () => {
		expect(render({ strokeDash: 'dash', lineCap: 'sq' })).toContain('d="M-3,0 L0,0 ');
		expect(render({ strokeDash: 'dash', lineCap: 'flat' })).not.toContain('d="M-3,0');
	});
});
