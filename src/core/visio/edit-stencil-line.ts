import { executableCellFormula } from './cell-formula';
import type { VisioRoutePoint } from './connector-route';
import type { End } from './edit-connector-layout';
import { inchTagged } from './edit-instance-checks';
import { recalculateInstanceCaches } from './edit-instance-recalculate';
import {
	instanceSheet,
	writeInstanceCell,
	effectiveNode,
	type InstanceCell,
	type InstanceSheet,
} from './edit-instance-sheet';
import {
	effectiveCells,
	effectiveNumber,
	isStencilConnector,
	isDynamicConnectorMaster,
	proveStencilConnector,
	stencilTemplate,
} from './edit-stencil-connector';
import { assertNoStencilDependents } from './edit-stencil-connector-layout';
import { transform } from './geometry';
import { fail, VisioPackageError } from './package-common';
import { attribute } from './sheet';
import { visioFormulaCachedValue } from './formula';

/**
 * Straight one-dimensional stencil shapes that are not Visio's Dynamic connector: the line and
 * arrow masters of the Connectors, Annotations and Arrows stencils. Their transform follows the
 * ends by the master's own formulas, in one of two forms:
 *
 * - a box: `Width=GUARD(EndX-BeginX)`, `Height=GUARD(ABS(EndY-BeginY))` and a flip flag, with the
 *   path running from the box's corner to the opposite one;
 * - a rotated run: `Width=SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)`, `Angle=ATAN2(EndY-BeginY,EndX-BeginX)`
 *   and a path along the run.
 *
 * Nothing here is keyed to those two forms: the instance is recalculated from its master's
 * formulas (`recalculateInstanceCaches`) and the result is accepted only if its transform puts
 * the path's two ends back on the connector's ends, so a master laid out any other way is
 * refused instead of being redrawn wrongly. The saved cells are the refreshed `F="Inh"` caches
 * Visio writes (recorded with `scripts/record-visio-stencil-line.ps1`).
 */
export const STENCIL_LINE_FORM =
	"This connector comes from a stencil and its ends do not move its drawing by the master's own formulas, so it cannot be rerouted here; it and the shapes glued to it cannot be moved or resized.";

const refuse = (): never => fail('UNSUPPORTED_GEOMETRY_EDIT', STENCIL_LINE_FORM);
const format = (value: number) => String(Object.is(value, -0) ? 0 : value);
const near = (a: number, b: number) =>
	Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
const NAMES = { begin: ['BeginX', 'BeginY'], end: ['EndX', 'EndY'] } as const;

/** A registered stencil instance of a 1-D master that is not Visio's Dynamic connector. */
export function isStencilLine(shape: Element | undefined): boolean {
	const template = stencilTemplate(shape);
	return !!shape && !!template && isStencilConnector(shape) && !isDynamicConnectorMaster(template);
}

/** Prove the instance can be recalculated from its master and follows its ends at all. */
export function proveStencilLine(shape: Element): InstanceSheet {
	const template = stencilTemplate(shape);
	if (!template) return refuse();
	let sheet: InstanceSheet;
	try {
		sheet = instanceSheet(shape, template);
	} catch (error) {
		if (error instanceof VisioPackageError && error.code === 'UNSUPPORTED_INSTANCE_EDIT')
			return refuse();
		throw error;
	}
	for (const name of ['pinx', 'piny', 'width', 'beginx', 'beginy', 'endx', 'endy'])
		if (!sheet.byName.has(name)) return refuse();
	// The transform has to be computed from the ends, not drawn once and left.
	for (const name of ['pinx', 'piny', 'width'] as const) {
		const cell = sheet.byName.get(name)!;
		const source = cell.local && attribute(cell.local, 'F') !== 'Inh' ? cell.local : cell.inherited;
		if (!executableCellFormula(attribute(source, 'F'))) return refuse();
	}
	return sheet;
}

/** The vertices of the shape's path in local coordinates, with every recalculated value applied. */
function pathVertices(
	sheet: InstanceSheet,
	width: number,
	height: number,
	writes: readonly { cell: InstanceCell; value: number }[],
): VisioRoutePoint[] {
	const rows = new Map<Element, { x?: number; y?: number; relative: boolean }>();
	for (const item of sheet.cells) {
		if (item.section?.name !== 'Geometry' || !item.row || (item.name !== 'X' && item.name !== 'Y'))
			continue;
		const written = writes.find((write) => write.cell === item);
		let number: number;
		if (written) number = written.value;
		else {
			const node = effectiveNode(item);
			if (!node || node.hasAttribute('E')) return refuse();
			try {
				number = visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).value;
			} catch {
				return refuse();
			}
		}
		const entry = rows.get(item.row) ?? { relative: item.relative };
		entry[item.name === 'X' ? 'x' : 'y'] = number;
		rows.set(item.row, entry);
	}
	return [...rows.values()]
		.filter((entry) => entry.x !== undefined && entry.y !== undefined)
		.map((entry) =>
			entry.relative ? { x: entry.x! * width, y: entry.y! * height } : { x: entry.x!, y: entry.y! },
		);
}

/** Lay a stencil line out between two points, the way Visio saves it after a glued shape moves. */
export function layoutStencilLine(
	root: Element,
	pageId: string,
	shape: Element,
	ends: Readonly<Record<End, VisioRoutePoint>>,
	check: () => void,
): readonly string[] {
	const sheet = proveStencilLine(shape);
	const effective = effectiveCells(shape);
	const cell = (name: string) => sheet.byName.get(name.toLowerCase());
	const number = (name: string, fallback: number): number => {
		const item = cell(name);
		const node = item && effectiveNode(item);
		if (!node) return fallback;
		if (node.hasAttribute('E')) return refuse();
		try {
			return visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).value;
		} catch {
			return refuse();
		}
	};
	const overrides = new Map<InstanceCell, number>();
	for (const end of ['begin', 'end'] as const) {
		const [x, y] = NAMES[end];
		const point = ends[end];
		if (
			(!near(number(x, point.x), point.x) || !near(number(y, point.y), point.y)) &&
			effectiveNumber(effective, end === 'begin' ? 'LockBegin' : 'LockEnd', 0) !== 0
		)
			fail('EDIT_PROTECTED_CELL', 'This connector is protected against moving its ends.');
		overrides.set(cell(x)!, point.x);
		overrides.set(cell(y)!, point.y);
	}
	assertNoStencilDependents(root, shape, check);
	const { writes } = recalculateInstanceCaches(sheet, overrides, check);
	const after = (name: string, fallback: number): number => {
		const written = writes.find((write) => write.cell === cell(name));
		return written ? written.value : number(name, fallback);
	};
	// The drawing's two ends have to land on the connector's ends, in either orientation of its box.
	const width = after('Width', 0),
		height = after('Height', 0);
	const matrix = transform(
		after('PinX', 0),
		after('PinY', 0),
		after('LocPinX', width / 2),
		after('LocPinY', height / 2),
		after('Angle', 0),
		after('FlipX', 0) !== 0,
		after('FlipY', 0) !== 0,
	);
	const at = (x: number, y: number) => ({
		x: matrix[0] * x + matrix[2] * y + matrix[4],
		y: matrix[1] * x + matrix[3] * y + matrix[5],
	});
	const lands = (from: VisioRoutePoint, to: VisioRoutePoint) =>
		near(from.x, ends.begin.x) &&
		near(from.y, ends.begin.y) &&
		near(to.x, ends.end.x) &&
		near(to.y, ends.end.y);
	if (
		![
			[at(0, 0), at(width, height)],
			[at(0, height / 2), at(width, height / 2)],
			[at(0, height), at(width, 0)],
		].some(([from, to]) => lands(from!, to!))
	)
		refuse();
	// And the drawn path has to reach them: a path of constants would be left where it was.
	const vertices = pathVertices(sheet, width, height, writes).map((point) => at(point.x, point.y));
	const reaches = (end: VisioRoutePoint) =>
		vertices.some((point) => near(point.x, end.x) && near(point.y, end.y));
	if (vertices.length < 2 || !reaches(ends.begin) || !reaches(ends.end)) refuse();
	// Apply: the ends keep their glue formulas and take the new values.
	for (const [item, value] of overrides) {
		const node = item.local;
		if (node) {
			node.setAttribute('V', format(value));
			node.removeAttribute('E');
		} else writeInstanceCell(sheet, item, format(value), { unit: 'IN' });
	}
	for (const write of writes) {
		check();
		const own = write.cell.local;
		if (own && executableCellFormula(attribute(own, 'F'))) {
			own.setAttribute('V', format(write.value));
			own.removeAttribute('E');
			continue;
		}
		const base = write.cell.inherited;
		const baseValue =
			base?.hasAttribute('V') && !base.hasAttribute('E') ? Number(attribute(base, 'V')) : NaN;
		if (own && attribute(own, 'F') === 'Inh' && near(baseValue, write.value)) {
			// Visio drops a refreshed cache that has come back to the master's own value.
			own.parentNode?.removeChild(own);
			continue;
		}
		if (!own && near(baseValue, write.value)) continue;
		writeInstanceCell(write.sheet, write.cell, format(write.value), {
			formula: 'Inh',
			...(inchTagged(write) ? { unit: 'IN' } : {}),
		});
	}
	return [pageId];
}

/** Prove a stencil connector of either kind: Visio's Dynamic connector, or a straight line master. */
export function proveAnyStencilConnector(shape: Element): void {
	if (isStencilLine(shape)) proveStencilLine(shape);
	else proveStencilConnector(shape);
}
