import { buildXml } from '../xml/index';
import type { VisioPackage } from './package';
import type { VisioDuplicateShapesEdit } from './edit-duplicate-commands';
import { attribute, children, VISIO_NS, VISIO_LEGACY_NS } from './sheet';
import { fail } from './package-common';
import { admitted, numeric, setCell } from './edit-geometry-admission';
import {
	assertShapeLocks,
	assertUnlayeredShape,
	formattingCell,
	assertEditableFormattingCell,
} from './edit-style-admission';
import { assertDuplicateScope } from './edit-duplicate-scope';
import { executableCellFormula } from './cell-formula';
import { mapVisioFormulaSyntax } from './formula-source';
import { analyzeVisioFormula } from './formula';
import { indexCells, key } from './edit-recalculate-index';
import { createVisioCellEvaluator } from './edit-recalculate-values';
import { recalculateVisioCells, type VisioCellKey } from './edit-recalculate';
import { createVisioDependencyQuery } from './edit-recalculate';
import { rerouteBlockedConnectors } from './edit-connector-follow';
import { assertGeometryPackageScope } from './edit-scope';
import { finishInstanceCopy, planInstanceCopy } from './edit-instance-duplicate';
import {
	assertConnectorCopyApart,
	assertGlueDroppable,
	finishConnectorCopy,
	finishLineCopy,
	isLineSheet,
	planLineCopy,
} from './edit-duplicate-line';
import { pageShapes, shapeIsStructural, stencilInstance } from './edit-instance-shape';
import { masterTemplate } from './edit-text-scope';

/** Native duplication drops UniqueID and gives each copy a fresh sheet name. */
function copyIdentity(shape: Element, newId: string, names: Map<string, Set<string>>): void {
	shape.setAttribute('ID', newId);
	shape.removeAttribute('UniqueID');
	for (const name of ['Name', 'NameU']) {
		const value = attribute(shape, name);
		if (value === undefined) continue;
		const occupied = names.get(name)!;
		let next = `${value.replace(/\.\d+$/, '')}.${newId}`;
		if (occupied.has(next.toLowerCase())) next = `Sheet.${newId}`;
		if (occupied.has(next.toLowerCase()))
			fail('UNSUPPORTED_DUPLICATE', 'Duplicated default shape name is already occupied.');
		if (next.length > 4096)
			fail('UNSUPPORTED_DUPLICATE', 'Duplicated name exceeds metadata limits.');
		shape.setAttribute(name, next);
		occupied.add(next.toLowerCase());
	}
}

/** Clone the source XML after every target, identity and affected-cache proof succeeds. */
export async function duplicateVisioShapes(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioDuplicateShapesEdit,
	check: () => void,
	detached?: { root: Element; pageId: string },
): Promise<readonly string[]> {
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const ids = new Set<string>();
	for (const node of Array.from(root.getElementsByTagName('*')))
		if (node.localName === 'Shape' && node.namespaceURI === root.namespaceURI) {
			const id = attribute(node, 'ID');
			if (!id || ids.has(id)) fail('INVALID_SHAPE_ID', 'Source shape IDs must be unique.');
			ids.add(id);
		}
	const newIds = new Set(edit.copies.map((copy) => copy.newShapeId));
	if ([...newIds].some((id) => ids.has(id)))
		fail('INVALID_SHAPE_ID', 'Duplicate shape ID already exists.');
	const sourceRoot = detached?.root ?? root;
	const sourcePageId = detached?.pageId ?? edit.pageId;
	const template = masterTemplate(pkg);
	// Copying a container or a list is the one case where container lookups elsewhere matter.
	let structural = false;
	for (const copy of edit.copies) {
		const source = pageShapes(sourceRoot).find((shape) => attribute(shape, 'ID') === copy.shapeId);
		structural ||= !source || (await shapeIsStructural(source, template));
	}
	await assertDuplicateScope(pkg, pagePaths, root, newIds, check, structural);
	// A glued target is copied alone, so its own cells are read as any other shape's.
	const glueTargets = new Set(
		children(sourceRoot, 'Connects').flatMap((container) =>
			children(container, 'Connect').map((row) => attribute(row, 'ToSheet')),
		),
	);
	const indexed = indexCells(detached ? new Map([[sourcePageId, sourceRoot]]) : roots, {
		check,
		scope: 'affected',
		glueShapes: new Set(
			pageShapes(sourceRoot).filter((shape) => glueTargets.has(attribute(shape, 'ID'))),
		),
	});
	const evaluate = createVisioCellEvaluator(indexed, { check }, true);
	const containers = children(root, 'Shapes');
	if (containers.length > 1) fail('INVALID_SHAPE_ID', 'Duplicate Shapes containers.');
	const container =
		containers[0] ?? root.ownerDocument!.createElementNS(root.namespaceURI, 'Shapes');
	const mapping = new Map(edit.copies.map((copy) => [copy.shapeId, copy.newShapeId]));
	const names = new Map(['Name', 'NameU'].map((name) => [name, new Set<string>()]));
	for (const node of Array.from(root.getElementsByTagName('*')))
		if (node.localName === 'Shape')
			for (const [name, occupied] of names) {
				const value = attribute(node, name);
				if (value !== undefined) occupied.add(value.toLowerCase());
			}
	const plans: {
		source: Element;
		id: string;
		x: number;
		y: number;
		template?: Element;
		connector?: boolean;
		line?: boolean;
	}[] = [];
	const copied = new Set(edit.copies.map((copy) => copy.shapeId));
	let cloneNodes = 0,
		cloneCharacters = buildXml(root).length;
	const size = (source: Element) => {
		cloneNodes += source.getElementsByTagName('*').length + 1;
		cloneCharacters += buildXml(source).length;
		if (cloneNodes > 100_000 || cloneCharacters > 16 * 1024 * 1024)
			fail('LIMIT_DUPLICATE', 'Duplicated XML exceeds bounded expansion limits.');
	};
	for (const copy of edit.copies) {
		check();
		// A stencil shape is copied as an instance of the same master, as Visio does.
		const stencil = await stencilInstance(
			sourceRoot,
			copy.shapeId,
			template,
			'UNSUPPORTED_DUPLICATE',
		);
		if (stencil) {
			assertGlueDroppable(sourceRoot, copy.shapeId, copied);
			const pin = await planInstanceCopy(
				pkg,
				edit.pageId,
				sourceRoot,
				stencil,
				edit.offsetX,
				edit.offsetY,
			);
			size(stencil.instance);
			plans.push({
				source: stencil.instance,
				id: copy.newShapeId,
				x: pin.x,
				y: pin.y,
				...(pin.connector ? { connector: true } : {}),
				template: stencil.template,
			});
			continue;
		}
		// A line is copied with its end points moved and its glue dropped, as Visio does.
		const line = pageShapes(sourceRoot).find((shape) => attribute(shape, 'ID') === copy.shapeId);
		if (line && !line.hasAttribute('Master') && isLineSheet(line)) {
			assertCloneLeaf(line);
			assertUnlayeredShape(line, document);
			assertShapeLocks(line, document, ['LockSelect']);
			assertGlueDroppable(sourceRoot, copy.shapeId, copied);
			const pin = planLineCopy(line);
			if (Math.abs(pin.x + edit.offsetX) > 1e6 || Math.abs(pin.y + edit.offsetY) > 1e6)
				fail('UNSUPPORTED_DUPLICATE', 'Duplicated pins exceed coordinate limits.');
			size(line);
			plans.push({ source: line, id: copy.newShapeId, x: pin.x, y: pin.y, line: true });
			continue;
		}
		const source = admitted(sourceRoot, copy.shapeId);
		assertCloneLeaf(source);
		assertUnlayeredShape(source, document);
		assertShapeLocks(source, document, ['LockSelect']);
		for (const connects of children(sourceRoot, 'Connects'))
			for (const connection of children(connects, 'Connect'))
				// A glued target is copied alone, as Visio does: its connectors stay with the original.
				if (attribute(connection, 'FromSheet') === copy.shapeId)
					fail('UNSUPPORTED_DUPLICATE', 'Glued shapes cannot be duplicated safely.');
		const x = numeric(formattingCell(source, 'PinX')),
			y = numeric(formattingCell(source, 'PinY'));
		for (const [name, offset] of [
			['PinX', edit.offsetX],
			['PinY', edit.offsetY],
		] as const) {
			if (offset !== 0) assertEditableFormattingCell(formattingCell(source, name));
			evaluate(key({ pageId: sourcePageId, shapeId: copy.shapeId, cell: name }));
		}
		if (Math.abs(x + edit.offsetX) > 1e6 || Math.abs(y + edit.offsetY) > 1e6)
			fail('UNSUPPORTED_DUPLICATE', 'Duplicated pins exceed coordinate limits.');
		size(source);
		plans.push({ source, id: copy.newShapeId, x: x + edit.offsetX, y: y + edit.offsetY });
	}
	const connectorSources = plans.filter((plan) => plan.connector).map((plan) => plan.source);
	if (connectorSources.length)
		assertConnectorCopyApart(root, connectorSources, edit.offsetX, edit.offsetY);
	const order = new Map(
		children(children(sourceRoot, 'Shapes')[0], 'Shape').map((source, index) => [source, index]),
	);
	plans.sort((a, b) => order.get(a.source)! - order.get(b.source)!);
	if (!container.parentNode) root.insertBefore(container, children(root, 'Connects')[0] ?? null);
	const changed: VisioCellKey[] = [];
	for (const plan of plans) {
		check();
		const shape = root.ownerDocument!.importNode(plan.source, true) as Element;
		copyIdentity(shape, plan.id, names);
		if (plan.line) finishLineCopy(shape, edit.offsetX, edit.offsetY);
		for (const node of [shape, ...Array.from(shape.getElementsByTagName('*'))]) {
			const source = executableCellFormula(attribute(node, 'F'));
			if (!source) continue;
			// A stencil shape's formulas may use functions the analyser does not know; a sheet
			// reference is still recognisable by its syntax.
			const named = plan.template
				? /\bSheet\.\d+!/i.test(source)
				: analyzeVisioFormula(source).references.some(
						(ref) => ref.shapeId !== undefined && mapping.has(ref.shapeId),
					);
			if (named)
				node.setAttribute(
					'F',
					mapVisioFormulaSyntax(source, (segment) =>
						segment.replace(/\bSheet\.(\d+)!/gi, (match, id: string) =>
							mapping.has(id) ? `Sheet.${mapping.get(id)}!` : match,
						),
					),
				);
		}
		if (plan.line) {
			container.appendChild(shape);
			continue;
		}
		if (plan.connector) {
			finishConnectorCopy(shape, plan.id, edit.offsetX, edit.offsetY);
			container.appendChild(shape);
			continue;
		}
		if (plan.template) {
			// Sub-shapes of a group instance take IDs beyond the page's and every copy's own.
			for (const id of newIds) ids.add(id);
			finishInstanceCopy(shape, plan.template, plan.x, plan.y, ids, check);
			container.appendChild(shape);
			continue;
		}
		for (const [name, value, offset] of [
			['PinX', plan.x, edit.offsetX],
			['PinY', plan.y, edit.offsetY],
		] as const)
			if (offset !== 0) {
				setCell(shape, name, value);
				changed.push({ pageId: edit.pageId, shapeId: plan.id, cell: name });
			}
		container.appendChild(shape);
	}
	const local = plans.filter((plan) => !plan.template && !plan.line);
	if (changed.length) {
		// Reuse inherited-style and package admission in the copies' actual sheet context.
		await assertGeometryPackageScope(
			pkg,
			pagePaths,
			local.map((plan) => ({
				type: 'move-shape',
				pageId: edit.pageId,
				shapeId: plan.id,
				x: plan.x,
				y: plan.y,
			})),
			check,
			roots,
		);
		const depends = createVisioDependencyQuery(roots, { check });
		for (const plan of local) {
			const clone = children(container, 'Shape').find((node) => attribute(node, 'ID') === plan.id)!;
			for (const section of children(clone, 'Section'))
				if (attribute(section, 'N') === 'Field')
					for (const row of children(section, 'Row'))
						for (const cell of children(row, 'Cell')) {
							check();
							const name = `Field.${attribute(row, 'N') ?? attribute(row, 'IX') ?? '0'}.${attribute(cell, 'N') ?? ''}`;
							if (depends({ pageId: edit.pageId, shapeId: plan.id, cell: name }, changed))
								fail(
									'UNSUPPORTED_DUPLICATE',
									'Pin-dependent text fields need native display-cache regeneration.',
								);
						}
		}
	}
	const dirty = new Set([edit.pageId]);
	for (const id of recalculateVisioCells(roots, changed, { check })) dirty.add(id);
	// Visio lays a right-angle connector out again around a shape copied onto its path.
	for (const plan of plans)
		if (!plan.line && !plan.connector)
			for (const id of rerouteBlockedConnectors(roots, edit.pageId, plan.id, check)) dirty.add(id);
	return [...dirty];
}

/** Unknown wrappers must not conceal additional sheet identities in a purported leaf copy. */
export function assertCloneLeaf(source: Element): void {
	if (
		Array.from(source.getElementsByTagName('*')).some(
			(node) =>
				node.localName === 'Shape' &&
				(node.namespaceURI === VISIO_NS || node.namespaceURI === VISIO_LEGACY_NS),
		)
	)
		fail('UNSUPPORTED_DUPLICATE', 'Nested Visio shape identities cannot be copied as leaf XML.');
}
