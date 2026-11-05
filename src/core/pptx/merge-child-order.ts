import { NS } from '../xml/namespaces';
import { FILL_CHOICE_ELEMENTS } from './core/core/runtime/fill-choice-group';
import {
	BLIP_FILL_ORDER,
	EFFECT_LST_ORDER,
	LN_CHILD_ORDER,
	SHAPE_STYLE_ORDER,
	SP_PR_ORDER,
	TC_PR_BORDERS_ORDER,
} from './core/utils/xml-reorder';

/** Shares a schema position between alternatives of one singleton choice. */
function sequence(order: readonly string[], choices: readonly (readonly string[])[] = []) {
	const ranks = new Map(order.map((name, index) => [name, index]));
	for (const choice of choices) {
		const positions = choice.flatMap((name) => {
			const position = ranks.get(name);
			return position === undefined ? [] : [position];
		});
		if (!positions.length) continue;
		const position = Math.min(...positions);
		for (const name of choice) if (ranks.has(name)) ranks.set(name, position);
	}
	return ranks;
}

const shapeChoices = [
	['a:custGeom', 'a:prstGeom'],
	FILL_CHOICE_ELEMENTS,
	['a:effectLst', 'a:effectDag'],
];
const shape = sequence(SP_PR_ORDER, shapeChoices);
const group = sequence(
	SP_PR_ORDER.filter((name) => !['a:custGeom', 'a:prstGeom', 'a:ln', 'a:sp3d'].includes(name)),
	shapeChoices,
);
const line = sequence(LN_CHILD_ORDER, [
	FILL_CHOICE_ELEMENTS,
	['a:prstDash', 'a:custDash'],
	['a:round', 'a:bevel', 'a:miter'],
]);
const blip = sequence(BLIP_FILL_ORDER, [['a:tile', 'a:stretch']]);
const effects = sequence(EFFECT_LST_ORDER);
const style = sequence(SHAPE_STYLE_ORDER);
const cell = sequence(TC_PR_BORDERS_ORDER, [FILL_CHOICE_ELEMENTS]);

/** Singleton DrawingML sequences used by slide and notes shapes; equal ranks are choices. */
export function pptxChildOrder(element: Element, parent: Element): number | undefined {
	if (element.namespaceURI !== NS.a) return undefined;
	let ranks: ReadonlyMap<string, number> | undefined;
	if (parent.namespaceURI === NS.p) {
		switch (parent.localName) {
			case 'spPr':
				ranks = shape;
				break;
			case 'grpSpPr':
				ranks = group;
				break;
			case 'blipFill':
				ranks = blip;
				break;
			case 'style':
				ranks = style;
				break;
		}
	} else if (parent.namespaceURI === NS.a) {
		switch (parent.localName) {
			case 'ln':
				ranks = line;
				break;
			case 'blipFill':
				ranks = blip;
				break;
			case 'effectLst':
				ranks = effects;
				break;
			case 'tcPr':
				ranks = cell;
				break;
		}
	}
	return ranks?.get('a:' + element.localName);
}
