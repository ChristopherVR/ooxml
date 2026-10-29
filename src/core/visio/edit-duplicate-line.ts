import { executableCellFormula } from './cell-formula';
import { isNativeGlueCell, visioGlueTrigger } from './edit-connector-glue';
import { cells, numeric } from './edit-geometry-admission';
import { proveLocalLine } from './edit-line-move';
import { instanceCell, type StencilInstance } from './edit-instance-shape';
import { fail } from './package-common';
import { attribute, children } from './sheet';

const ENDS = ['BeginX', 'BeginY', 'EndX', 'EndY'] as const;
const connects = (root: Element) =>
	children(root, 'Connects').flatMap((container) => children(container, 'Connect'));

/** Sheets the connector `id` is glued to, by the page's Connect rows. */
export const gluedTargets = (root: Element, id: string): string[] =>
	connects(root)
		.filter((row) => attribute(row, 'FromSheet') === id)
		.map((row) => attribute(row, 'ToSheet') ?? '');

/**
 * Visio copies a connector that is glued to shapes without its glue (evidence:
 * native-duplicate-glue-2026-10-29.json): the copy keeps the end points and has no Connect rows. A
 * connector copied together with a shape it is glued to would keep its glue to the copy, which is
 * not modelled, so that selection is refused.
 */
export function assertGlueDroppable(
	root: Element,
	id: string,
	copied: ReadonlySet<string>,
	code = 'UNSUPPORTED_DUPLICATE',
): void {
	if (gluedTargets(root, id).some((target) => copied.has(target)))
		fail(code, 'A glued connector cannot be copied together with a shape it is glued to yet.');
}

/** The Begin/End cells of a line sheet may hold nothing or a native glue formula, never another. */
function assertEndFormulas(read: (name: string) => Element | undefined, code: string): void {
	for (const name of ENDS) {
		const formula = executableCellFormula(attribute(read(name), 'F'));
		if (formula && !isNativeGlueCell(name, formula))
			fail(code, `${name} of the line holds a formula that is not native glue.`);
	}
}

/** Whether a local sheet is a line: it has end point cells. */
export const isLineSheet = (shape: Element): boolean => ENDS.some((name) => cells(shape).has(name));

/** Admit a local straight line (glued or not) as the source of a copy and return its pin. */
export function planLineCopy(shape: Element, code = 'UNSUPPORTED_DUPLICATE') {
	proveLocalLine(shape);
	assertEndFormulas((name) => cells(shape).get(name), code);
	const local = cells(shape);
	return { x: numeric(local.get('PinX')), y: numeric(local.get('PinY')) };
}

const shift = (node: Element | undefined, delta: number, name: string): void => {
	if (!node) fail('UNSUPPORTED_DUPLICATE', `The line has no local ${name}.`);
	const value = numeric(node);
	node!.setAttribute('V', String(Object.is(value + delta, -0) ? 0 : value + delta));
	node!.removeAttribute('E');
};

/**
 * Move a copied line by (`dx`, `dy`) as Visio does: every end point becomes a plain value (the
 * glue formula goes), and the midpoint, which follows the ends, moves with them.
 */
export function finishLineCopy(shape: Element, dx: number, dy: number): void {
	const local = cells(shape);
	for (const name of ENDS) {
		const node = local.get(name);
		const formula = executableCellFormula(attribute(node, 'F'));
		if (formula && isNativeGlueCell(name, formula)) node!.removeAttribute('F');
		shift(node, name.endsWith('X') ? dx : dy, name);
	}
	shift(local.get('PinX'), dx, 'PinX');
	shift(local.get('PinY'), dy, 'PinY');
}

/** A stencil connector (Visio's Dynamic connector) is a 1D instance with end point cells. */
export const isStencilConnector = (target: StencilInstance): boolean =>
	ENDS.some((name) => instanceCell(target, name));

/** Admit a stencil connector whose end points, pin and triggers are local cells. */
export function planConnectorCopy(
	target: StencilInstance,
	code = 'UNSUPPORTED_DUPLICATE',
): { x: number; y: number } {
	const local = cells(target.instance);
	for (const name of [...ENDS, 'PinX', 'PinY']) {
		const node = local.get(name);
		if (!node?.hasAttribute('V') || node.hasAttribute('E'))
			fail(code, `The stencil connector has no usable local ${name}.`);
		numeric(node);
	}
	assertEndFormulas((name) => local.get(name), code);
	return { x: numeric(local.get('PinX')), y: numeric(local.get('PinY')) };
}

/** Copy of a stencil connector: ends and pin move, and the triggers name the copy's own sheet. */
export function finishConnectorCopy(shape: Element, id: string, dx: number, dy: number): void {
	const local = cells(shape);
	for (const name of ENDS) shift(local.get(name), name.endsWith('X') ? dx : dy, name);
	shift(local.get('PinX'), dx, 'PinX');
	shift(local.get('PinY'), dy, 'PinY');
	// Visio rewrites the triggers of an unglued connector to its own sheet.
	for (const name of ['BegTrigger', 'EndTrigger']) {
		const node = local.get(name);
		if (!node) continue;
		node.setAttribute('F', visioGlueTrigger(id));
		node.setAttribute('V', '1');
	}
}

/** The inflation of a connector's end point box: Visio keeps its routes about a clearance apart. */
const CLEARANCE = 0.25;
const endsBox = (shape: Element, dx: number, dy: number) => {
	const local = cells(shape);
	const x = [numeric(local.get('BeginX')) + dx, numeric(local.get('EndX')) + dx];
	const y = [numeric(local.get('BeginY')) + dy, numeric(local.get('EndY')) + dy];
	return {
		minX: Math.min(...x) - CLEARANCE,
		maxX: Math.max(...x) + CLEARANCE,
		minY: Math.min(...y) - CLEARANCE,
		maxY: Math.max(...y) + CLEARANCE,
	};
};

/**
 * Visio lays out a connector that crosses or touches another dynamic connector again, which moves
 * both routes a little on the next layout (native-edit-families-full-2026-10-29.json, the
 * overlapping copies of `glue-duplicate-line`). The copy of a stencil connector is therefore
 * refused when its end point box meets the box of a dynamically glued connector it does not
 * replace, instead of leaving routes in the file that Visio would change.
 */
export function assertConnectorCopyApart(
	root: Element,
	copies: readonly Element[],
	dx: number,
	dy: number,
	code = 'UNSUPPORTED_DUPLICATE',
): void {
	const boxes = copies.map((shape) => endsBox(shape, dx, dy));
	for (const container of children(root, 'Shapes'))
		for (const other of children(container, 'Shape')) {
			const local = cells(other);
			const formula = executableCellFormula(attribute(local.get('BeginX'), 'F')) ?? '';
			if (!/^_WALKGLUE\(/i.test(formula)) continue;
			const box = endsBox(other, 0, 0);
			for (const copy of boxes)
				if (
					copy.minX <= box.maxX &&
					box.minX <= copy.maxX &&
					copy.minY <= box.maxY &&
					box.minY <= copy.maxY
				)
					fail(
						code,
						'The copy of a connector would meet another connector, which Visio lays out again.',
					);
		}
}
