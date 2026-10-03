import { parseVsdx } from '../visio/parser.js';
import type { VisioShape } from '../visio/model.js';

/** JSON-safe scene inspection without embedding binary images in tool responses. */
export interface VisioShapeSummary {
	id: string;
	name: string;
	kind: VisioShape['kind'];
	width: number;
	height: number;
	transform: VisioShape['transform'];
	text: string;
	hidden: boolean;
	shapeData: VisioShape['shapeData'];
	hyperlinks: VisioShape['hyperlinks'];
	children: VisioShapeSummary[];
}

function inspectShape(shape: VisioShape): VisioShapeSummary {
	return {
		id: shape.id,
		name: shape.name,
		kind: shape.kind,
		width: shape.width,
		height: shape.height,
		transform: shape.transform,
		text: shape.text.plainText,
		hidden: shape.hidden,
		shapeData: shape.shapeData,
		hyperlinks: shape.hyperlinks,
		children: shape.children.map(inspectShape),
	};
}

export async function inspectVisio(bytes: Uint8Array) {
	const document = await parseVsdx(bytes);
	return {
		pages: document.pages.map((page) => ({
			id: page.id,
			name: page.name,
			width: page.width,
			height: page.height,
			isBackground: page.isBackground,
			backgroundPageId: page.backgroundPageId,
			shapes: page.shapes.map(inspectShape),
			connectors: page.connectors,
			layers: page.layers,
		})),
		diagnostics: document.diagnostics,
	};
}
