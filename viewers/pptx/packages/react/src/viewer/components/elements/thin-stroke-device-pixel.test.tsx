import { screenStrokeWidth } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { getShapeVisualStyle } from '../../utils/shape-visual-style';
import { TableCellDiagonalBorders } from '../../utils/table-diagonal-borders';
import { ConnectorElementRenderer } from './ConnectorElementRenderer';
import type { ConnectorRendererProps } from './element-renderer-types';
import { shapeParams } from './element-shape-params';
import { ShapeEffectOverlay } from './ShapeEffectOverlay';

/**
 * Issue #23: a stroke never paints thinner than one device pixel on screen.
 * The stage publishes `--pptx-device-px`; every painted width is
 * `max(<authored>px, var(--pptx-device-px, 0px))`, with the authored width kept
 * as the attribute for anything that reads attributes. Shared helpers decide
 * the value; these pin the React wiring of each surface.
 */
const HAIRLINE = 0.1 * (96 / 72);
const STEP = 'max(var(--pptx-device-px, 0px), 0.001px)';
const screen = (px: string) => `max(${STEP}, round(${px}px, ${STEP}))`;

function line(): PptxElement {
	return {
		id: 'line-1',
		type: 'shape',
		shapeType: 'line',
		x: 0,
		y: 0,
		width: 400,
		height: 0,
		shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
	} as unknown as PptxElement;
}

describe('react: thin strokes stay one device pixel on screen', () => {
	it('paints a hairline `line` preset outline at max(authored, one device pixel)', () => {
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={line()} />);
		expect(html).toContain(`stroke-width="${HAIRLINE}"`);
		expect(html).toContain(`stroke-width:${screen('0.1333')}`);
	});

	it('holds an inset CSS border at one device pixel and leaves a no-line shape at 0', () => {
		const inset = {
			...line(),
			shapeType: 'rect',
			height: 50,
			shapeStyle: { strokeColor: '#000000', strokeWidth: 0.5, lineAlignment: 'in' },
		} as unknown as PptxElement;
		const insetParams = shapeParams(inset);
		expect(
			getShapeVisualStyle(inset, insetParams.hf, insetParams.fc, insetParams.sw, insetParams.sc)
				.borderWidth,
		).toBe('max(0.5px, var(--pptx-device-px, 0px))');
		const bare = {
			...inset,
			shapeStyle: { fillColor: '#ff0000' },
		} as unknown as PptxElement;
		const bareParams = shapeParams(bare);
		expect(
			getShapeVisualStyle(bare, bareParams.hf, bareParams.fc, bareParams.sw, bareParams.sc)
				.borderWidth,
		).toBe(0);
	});

	it('holds a connector strand at one device pixel', () => {
		const el = {
			id: 'conn-1',
			type: 'connector',
			shapeType: 'straightConnector1',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			shapeStyle: { strokeColor: '#000000', strokeWidth: HAIRLINE },
		} as unknown as PptxElement;
		const props = {
			el,
			isSelected: false,
			canInteract: false,
			showResizeHandles: false,
			showHoverBorder: false,
			selectionColorClass: 'blue-500',
			opacity: 1,
			zIndex: 0,
		} as unknown as ConnectorRendererProps;
		const html = renderToStaticMarkup(<ConnectorElementRenderer {...props} />);
		expect(html).toContain(`stroke-width:${screenStrokeWidth(HAIRLINE, 1)}`);
	});

	it('holds a table diagonal border at one device pixel', () => {
		const html = renderToStaticMarkup(
			<TableCellDiagonalBorders diag={{ diagDownColor: '#000000', diagDownWidth: 0.5 }} />,
		);
		expect(html).toContain('stroke-width="0.5"');
		expect(html).toContain(`stroke-width:${screen('0.5')}`);
	});

	it('snaps a horizontal connector to the device grid, but not a diagonal one', () => {
		const render = (height: number) => {
			const el = {
				id: 'conn-2',
				type: 'connector',
				shapeType: 'straightConnector1',
				x: 27,
				y: 27,
				width: 400,
				height,
				shapeStyle: { strokeColor: '#000000', strokeWidth: 1.3333 },
			} as unknown as PptxElement;
			const props = {
				el,
				isSelected: false,
				canInteract: false,
				showResizeHandles: false,
				showHoverBorder: false,
				selectionColorClass: 'blue-500',
				opacity: 1,
				zIndex: 0,
			} as unknown as ConnectorRendererProps;
			return renderToStaticMarkup(<ConnectorElementRenderer {...props} />);
		};
		expect(render(0)).toContain('transform:translate(0px, calc(round(27px');
		expect(render(40)).not.toContain('calc(round(27px');
	});

	it('snaps every edge of a plain rectangle outline, from its exact EMU position', () => {
		const rect = {
			id: 'rect-1',
			type: 'shape',
			shapeType: 'rect',
			x: 454,
			y: 387,
			width: 80,
			height: 67,
			xEmu: 4321810,
			yEmu: 3688080,
			widthEmu: 762000,
			heightEmu: 641350,
			shapeStyle: { strokeColor: '#000000', strokeWidth: 1 },
		} as unknown as PptxElement;
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={rect} />);
		expect(html).toContain('scale(calc((80px');
		expect(html).toContain('(67.3333px');
	});

	it('ends a connector with a flat cap unless the line declares another', () => {
		const render = (shapeStyle: Record<string, unknown>) => {
			const el = {
				id: 'conn-3',
				type: 'connector',
				shapeType: 'straightConnector1',
				x: 27,
				y: 27,
				width: 400,
				height: 0,
				shapeStyle,
			} as unknown as PptxElement;
			const props = {
				el,
				isSelected: false,
				canInteract: false,
				showResizeHandles: false,
				showHoverBorder: false,
				selectionColorClass: 'blue-500',
				opacity: 1,
				zIndex: 0,
			} as unknown as ConnectorRendererProps;
			const html = renderToStaticMarkup(<ConnectorElementRenderer {...props} />);
			return /<(?:line|path)[^>]*data-pptx-line[^>]*>/u.exec(html)?.[0] ?? '';
		};
		expect(render({ strokeColor: '#000000', strokeWidth: 12 })).toContain('stroke-linecap="butt"');
		expect(render({ strokeColor: '#000000', strokeWidth: 12, lineCap: 'rnd' })).toContain(
			'stroke-linecap="round"',
		);
	});
});
