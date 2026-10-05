import type { VisioDocument, VisioShape } from 'ooxml-core/visio';

function box(
	id: string,
	name: string,
	x: number,
	y: number,
	width: number,
	height: number,
	fill: string,
	kind: 'box' | 'diamond' = 'box',
): VisioShape {
	const path =
		kind === 'diamond'
			? `M ${width / 2} 0 L ${width} ${height / 2} L ${width / 2} ${height} L 0 ${height / 2} Z`
			: `M 0 0 L ${width} 0 L ${width} ${height} L 0 ${height} Z`;
	return {
		id,
		name,
		kind: 'shape',
		width,
		height,
		transform: [1, 0, 0, 1, x, y],
		hidden: false,
		children: [],
		geometry: [{ path, fill: true, stroke: true }],
		style: {
			fill,
			lineColor: '#326765',
			lineWidth: 0.015,
			linePattern: 1,
			fillOpacity: 1,
			lineOpacity: 1,
			startArrow: 0,
			endArrow: 0,
		},
		text: {
			plainText: name,
			runs: [],
			fontFamily: 'Arial',
			fontSize: 0.15,
			color: '#203d40',
			horizontalAlign: 'center',
			verticalAlign: 'middle',
			transform: [1, 0, 0, 1, 0, 0],
			width,
			height,
			margins: { left: 0.06, right: 0.06, top: 0.06, bottom: 0.06 },
		},
	};
}
function line(id: string, x: number, y: number, path: string): VisioShape {
	const shape = box(id, '', x, y, 0, 0, 'none');
	shape.kind = 'connector';
	shape.geometry = [{ path, fill: false, stroke: true }];
	return shape;
}
/** Original sample scene, not a Visio-authored reference fixture. */
export const demoDocument: VisioDocument = {
	format: 'vsdx',
	diagnostics: [],
	pages: [
		{
			id: '1',
			name: 'Release workflow',
			width: 8.5,
			height: 7,
			isBackground: false,
			connectors: [],
			shapes: [
				box('s1', 'Start with an idea', 2.8, 5.5, 2.8, 0.7, '#d8efeb'),
				line('c1', 4.2, 5.5, 'M 0 0 L 0 -.5'),
				box('s2', 'Design and build', 2.8, 4.3, 2.8, 0.7, '#f3eee1'),
				line('c2', 4.2, 4.3, 'M 0 0 L 0 -.45'),
				box('s3', 'Ready?', 3.35, 2.8, 1.7, 1.05, '#dfebf7', 'diamond'),
				line('c3', 4.2, 2.8, 'M 0 0 L 0 -.5'),
				box('s4', 'Review and release', 2.8, 1.6, 2.8, 0.7, '#d8efeb'),
				line('c4', 5.05, 3.325, 'M 0 0 L 1.3 0 L 1.3 1.325 L .55 1.325'),
				box(
					'note',
					'Explore pages · Select shapes\nZoom with + / − · Fit with 0',
					2.3,
					0.45,
					3.8,
					0.7,
					'#ffffff',
				),
			],
		},
		{
			id: '2',
			name: 'Architecture',
			width: 8.5,
			height: 5.5,
			isBackground: false,
			connectors: [],
			shapes: [
				box('a1', 'Your framework', 0.7, 3.5, 2, 0.9, '#dfebf7'),
				box('a2', 'One viewer', 3.2, 3.5, 2, 0.9, '#d8efeb'),
				box('a3', 'ooxml-core', 5.7, 3.5, 2, 0.9, '#f3eee1'),
				line('ac1', 2.7, 3.95, 'M 0 0 L .5 0'),
				line('ac2', 5.2, 3.95, 'M 0 0 L .5 0'),
				box(
					'a4',
					'Framework-neutral rendering\nLocal files. Explicit compatibility notes.',
					1.55,
					1.5,
					5.4,
					1,
					'#ffffff',
				),
			],
		},
	],
};
