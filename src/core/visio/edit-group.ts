import type { VisioGroupShapesEdit, VisioUngroupShapeEdit } from './edit-group-commands';
import { attribute, children, VISIO_NS, VISIO_LEGACY_NS } from './sheet';
import { fail } from './package-common';
import { cells, numeric, setCell, editableCell, protectedShape } from './edit-geometry-admission';
import { assertShapeLocks, assertUnlayeredShape } from './edit-style-admission';
import { applyShapeCreationStyles } from './edit-shape-create';
import { proveLocalShapeTree, assertUngluedTree } from './edit-group-rotation';
import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula } from './formula';
import type { VisioFormulaAnalysis } from './formula-evaluate';
import { transform } from './geometry';

const isSheet = (node: Node | null): node is Element =>
	!!node &&
	node.nodeType === 1 &&
	(node as Element).localName === 'Shape' &&
	((node as Element).namespaceURI === VISIO_NS ||
		(node as Element).namespaceURI === VISIO_LEGACY_NS);
const clean = (value: number) => (Math.abs(value) < 1e-12 ? 0 : value);

/** Visit every executable formula on the page with its owning sheet. Unknown scope fails. */
function eachFormula(
	root: Element,
	check: () => void,
	visit: (node: Element, owner: Element | undefined, analysis: VisioFormulaAnalysis) => void,
): void {
	for (const node of Array.from(root.getElementsByTagName('*'))) {
		check();
		const source = executableCellFormula(attribute(node, 'F'));
		if (!source) continue;
		let analysis: VisioFormulaAnalysis;
		try {
			analysis = analyzeVisioFormula(source, { onStep: check });
		} catch {
			fail('UNSUPPORTED_GROUP_EDIT', 'Page formulas cannot be proven independent of grouping.');
		}
		if (analysis.dynamic)
			fail('UNSUPPORTED_GROUP_EDIT', 'Dynamic formulas may depend on the sheet hierarchy.');
		let owner: Node | null = node.parentNode;
		while (owner && !isSheet(owner)) owner = owner.parentNode;
		visit(node, owner ?? undefined, analysis);
	}
}
function topLevel(root: Element): Element[] {
	const containers = children(root, 'Shapes');
	if (containers.length !== 1)
		fail('UNSUPPORTED_GROUP_EDIT', 'One local Shapes container is required.');
	return children(containers[0], 'Shape');
}
function ownerId(owner: Element | undefined): string | undefined {
	return owner ? attribute(owner, 'ID') : undefined;
}
/** Shared sheet admission: locks (including LockGroup) and the ordinary layer scope. */
function assertGroupable(shape: Element, document: Element): void {
	protectedShape(shape, document);
	assertShapeLocks(shape, document, ['LockGroup']);
	assertUnlayeredShape(shape, document);
}
const transformOf = (shape: Element) => {
	const local = cells(shape);
	const width = numeric(local.get('Width')),
		height = numeric(local.get('Height'));
	return {
		width,
		height,
		pinX: numeric(local.get('PinX'), width / 2),
		pinY: numeric(local.get('PinY'), height / 2),
		angle: numeric(local.get('Angle'), 0),
		flipX: numeric(local.get('FlipX'), 0) !== 0,
		flipY: numeric(local.get('FlipY'), 0) !== 0,
		matrix: transform(
			numeric(local.get('PinX'), width / 2),
			numeric(local.get('PinY'), height / 2),
			numeric(local.get('LocPinX'), width / 2),
			numeric(local.get('LocPinY'), height / 2),
			numeric(local.get('Angle'), 0),
			numeric(local.get('FlipX'), 0) !== 0,
			numeric(local.get('FlipY'), 0) !== 0,
		),
	};
};

/** The page-space bounds of top-level sheets' alignment boxes (through pin, rotation and flips). */
export function visioShapeBounds(shapes: readonly Element[]): {
	minX: number;
	minY: number;
	maxX: number;
	maxY: number;
} {
	let minX = Infinity,
		minY = Infinity,
		maxX = -Infinity,
		maxY = -Infinity;
	for (const shape of shapes) {
		const box = transformOf(shape);
		const [a, b, c, d, e, f] = box.matrix;
		for (const [x, y] of [
			[0, 0],
			[box.width, 0],
			[0, box.height],
			[box.width, box.height],
		] as const) {
			const px = a * x + c * y + e,
				py = b * x + d * y + f;
			minX = Math.min(minX, px);
			maxX = Math.max(maxX, px);
			minY = Math.min(minY, py);
			maxY = Math.max(maxY, py);
		}
	}
	return { minX, minY, maxX, maxY };
}

/** Native Group: the new sheet bounds the members' alignment boxes and owns them in its
 * Shapes container. Member pins move into group-local inches; every other cell is preserved.
 */
export function groupVisioShapes(
	root: Element,
	document: Element,
	edit: VisioGroupShapesEdit,
	check: () => void,
): boolean {
	check();
	const siblings = topLevel(root);
	for (const node of Array.from(root.getElementsByTagName('*')))
		if (isSheet(node) && attribute(node, 'ID') === edit.shapeId)
			fail('INVALID_SHAPE_ID', 'The new group ID already exists on the page.');
	const wanted = new Set(edit.memberIds);
	const members = siblings.filter((shape) => wanted.has(attribute(shape, 'ID') ?? ''));
	if (members.length !== wanted.size)
		fail('EDIT_TARGET_NOT_FOUND', 'Group members must be unique top-level local shapes.');
	const memberIds = new Set<string>(),
		subtree = new Set<string>();
	for (const member of members) {
		for (const id of proveLocalShapeTree(member, check)) subtree.add(id);
		memberIds.add(attribute(member, 'ID')!);
		assertGroupable(member, document);
		for (const name of ['PinX', 'PinY']) editableCell(cells(member).get(name));
	}
	assertUngluedTree(root, subtree, 'Glued shapes cannot be grouped without connection updates.');
	eachFormula(root, check, (_node, owner, analysis) => {
		const from = ownerId(owner);
		for (const ref of analysis.references) {
			const target = ref.shapeId ?? from;
			if (target && memberIds.has(target) && /^pin[xy]$/i.test(ref.cell))
				fail('UNSUPPORTED_GROUP_EDIT', 'Formulas depend on a member position.');
			if (!ref.shapeId && ref.cell.includes('!') && from && subtree.has(from))
				fail('UNSUPPORTED_GROUP_EDIT', 'Member formulas reference page or document sheets.');
		}
	});
	const { minX, minY, maxX, maxY } = visioShapeBounds(members);
	const plans = members.map((member) => {
		const box = transformOf(member);
		return { member, pinX: box.pinX, pinY: box.pinY };
	});
	const width = maxX - minX,
		height = maxY - minY;
	if (!(width > 0 && height > 0) || ![minX, minY, maxX, maxY].every((v) => Math.abs(v) <= 1e6))
		fail('UNSUPPORTED_GROUP_EDIT', 'Group bounds must be positive and within coordinate limits.');
	// No mutation occurs until every member, lock, glue and formula proof has succeeded.
	check();
	const doc = root.ownerDocument!;
	const group = doc.createElementNS(root.namespaceURI, 'Shape');
	group.setAttribute('ID', edit.shapeId);
	group.setAttribute('Type', 'Group');
	applyShapeCreationStyles(group, document);
	for (const [name, value, formula] of [
		['PinX', minX + width / 2],
		['PinY', minY + height / 2],
		['Width', width],
		['Height', height],
		['LocPinX', width / 2, 'Width*0.5'],
		['LocPinY', height / 2, 'Height*0.5'],
		['Angle', 0],
		['FlipX', 0],
		['FlipY', 0],
		['ResizeMode', 0],
	] as const)
		setCell(group, name, value, formula);
	const inner = doc.createElementNS(root.namespaceURI, 'Shapes');
	group.appendChild(inner);
	// Visio places the new group at the stacking position of its topmost member.
	const topmost = members[members.length - 1]!;
	topmost.parentNode!.insertBefore(group, topmost);
	for (const plan of plans) {
		setCell(plan.member, 'PinX', clean(plan.pinX - minX));
		setCell(plan.member, 'PinY', clean(plan.pinY - minY));
		inner.appendChild(plan.member);
	}
	return true;
}

/** Native Ungroup: direct members return to page coordinates through the group's pin,
 * rotation and flips; the group sheet (with its own data) is removed. Returns member IDs.
 */
export function ungroupVisioShape(
	root: Element,
	document: Element,
	edit: VisioUngroupShapeEdit,
	check: () => void,
): readonly string[] {
	check();
	const candidates = topLevel(root).filter((shape) => attribute(shape, 'ID') === edit.shapeId);
	if (candidates.length !== 1 || attribute(candidates[0], 'Type') !== 'Group')
		fail('EDIT_TARGET_NOT_FOUND', 'A unique top-level local group is required.');
	const group = candidates[0]!;
	const members = children(children(group, 'Shapes')[0], 'Shape');
	if (!members.length) fail('UNSUPPORTED_GROUP_EDIT', 'The group has no member shapes.');
	const subtree = proveLocalShapeTree(group, check);
	assertGroupable(group, document);
	if (
		children(group, 'Section').some((section) => attribute(section, 'N') === 'Geometry') ||
		children(group, 'Text').some((text) => (text.textContent ?? '').trim())
	)
		fail('UNSUPPORTED_GROUP_EDIT', 'Groups with their own geometry or text cannot be ungrouped.');
	assertUngluedTree(root, subtree, 'Glued shapes cannot be ungrouped without connection updates.');
	const memberIds = new Set(members.map((member) => attribute(member, 'ID')!));
	const changing = /^(pinx|piny|angle|flipx|flipy)$/i;
	// Native members keep group-scaling formulas such as Width = Sheet.5!Width*0.5. Ungrouping
	// converts those (and only those) to their cached values, as Visio does.
	const scaling = new Set<Element>();
	eachFormula(root, check, (node, owner, analysis) => {
		const from = ownerId(owner);
		const direct =
			owner && memberIds.has(from ?? '') && node.parentNode === owner && node.localName === 'Cell';
		const groupScale =
			direct &&
			/^(PinX|PinY|Width|Height)$/.test(attribute(node, 'N') ?? '') &&
			!analysis.guarded &&
			!analysis.unsupportedFunctions.length &&
			analysis.references.length > 0 &&
			analysis.references.every(
				(ref) => ref.shapeId === edit.shapeId && /^(width|height)$/i.test(ref.cell),
			);
		if (groupScale) {
			scaling.add(node);
			return;
		}
		if (direct && changing.test(attribute(node, 'N') ?? ''))
			fail('UNSUPPORTED_GROUP_EDIT', 'Member transform formulas cannot be ungrouped safely.');
		for (const ref of analysis.references) {
			const target = ref.shapeId ?? from;
			if (ref.shapeId === edit.shapeId)
				fail('UNSUPPORTED_GROUP_EDIT', 'Formulas reference the group sheet being removed.');
			if (target && memberIds.has(target) && changing.test(ref.cell))
				fail('UNSUPPORTED_GROUP_EDIT', 'Formulas depend on a member position or rotation.');
		}
	});
	for (const member of members) {
		const local = cells(member);
		for (const name of ['PinX', 'PinY', 'Angle', 'FlipX', 'FlipY']) {
			const cell = local.get(name);
			if (!cell || !scaling.has(cell)) editableCell(cell);
		}
		protectedShape(member, document);
	}
	const parent = transformOf(group);
	const [a, b, c, d, e, f] = parent.matrix;
	const oneFlip = parent.flipX !== parent.flipY;
	const plans = members.map((member) => {
		const box = transformOf(member);
		let angle = parent.angle + (oneFlip ? -box.angle : box.angle);
		angle = Math.atan2(Math.sin(angle), Math.cos(angle));
		if (Math.abs(Math.abs(angle) - Math.PI) < 1e-12) angle = Math.PI;
		const plan = {
			member,
			pinX: clean(a * box.pinX + c * box.pinY + e),
			pinY: clean(b * box.pinX + d * box.pinY + f),
			angle: clean(angle),
			flipX: box.flipX !== parent.flipX ? 1 : 0,
			flipY: box.flipY !== parent.flipY ? 1 : 0,
			box,
		};
		if (![plan.pinX, plan.pinY].every((value) => Math.abs(value) <= 1e6))
			fail('UNSUPPORTED_GROUP_EDIT', 'Ungrouped pins exceed coordinate limits.');
		return plan;
	});
	check();
	for (const node of scaling) node.removeAttribute('F');
	const container = group.parentNode!;
	for (const plan of plans) {
		setCell(plan.member, 'PinX', plan.pinX);
		setCell(plan.member, 'PinY', plan.pinY);
		if (plan.angle !== plan.box.angle || cells(plan.member).has('Angle'))
			setCell(plan.member, 'Angle', plan.angle);
		for (const [name, value, before] of [
			['FlipX', plan.flipX, plan.box.flipX],
			['FlipY', plan.flipY, plan.box.flipY],
		] as const)
			if (value !== (before ? 1 : 0)) setCell(plan.member, name, value);
		container.insertBefore(plan.member, group);
	}
	container.removeChild(group);
	return [...memberIds];
}
