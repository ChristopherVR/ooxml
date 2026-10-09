import type { VisioPackage } from './package';
import type { VisioFormatEdit } from './edit-formatting-commands';
import { attribute, children } from './sheet';
import { fail } from './package-common';
import { visioFormulaCachedValue } from './formula';
import { textFormattingWrites } from './edit-formatting-text';
import { shapeFormattingPlan, assertShapeFormattingPaintScope } from './edit-formatting-paint';
import { assertFormattingDependencies } from './edit-formatting-scope';
import {
	assertEditableFormattingCell,
	assertShapeLocks,
	assertUnlayeredShape,
	effectiveShapeCell,
	formattingCell,
	type FormattingCategory,
	type FormattingRowContext,
} from './edit-style-admission';

function targetShape(root: Element, shapeId: string): Element {
	const containers = children(root, 'Shapes');
	if (containers.length !== 1)
		fail('UNSUPPORTED_FORMAT_EDIT', 'One local Shapes container is required.');
	const candidates = children(containers[0], 'Shape').filter(
		(shape) => attribute(shape, 'ID') === shapeId,
	);
	if (candidates.length !== 1)
		fail('EDIT_TARGET_NOT_FOUND', 'A unique top-level local shape is required.');
	const shape = candidates[0]!;
	if (
		shape.hasAttribute('Master') ||
		shape.hasAttribute('MasterShape') ||
		shape.hasAttribute('Del') ||
		children(shape, 'Shapes').length ||
		children(shape, 'ForeignData').length ||
		children(shape, 'Rel').length ||
		!['Shape', undefined].includes(attribute(shape, 'Type'))
	)
		fail(
			'UNSUPPORTED_FORMAT_EDIT',
			'Master, group, foreign, deleted and non-shape formatting is unsupported.',
		);
	return shape;
}
export interface FormattingWrite {
	name: string;
	value: string;
	unit?: string;
	formula?: string;
	/** Remove a stored unit: the value is a plain scalar (negative SpLine multiples). */
	dropUnit?: boolean;
	category: FormattingCategory;
}

export async function applyFormattingEdit(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioFormatEdit,
	check: () => void,
): Promise<boolean> {
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const shape = targetShape(root, edit.shapeId);
	assertUnlayeredShape(shape, document);
	assertShapeLocks(
		shape,
		document,
		edit.type === 'format-text' ? ['LockFormat', 'LockTextEdit'] : ['LockFormat'],
	);
	let writes: FormattingWrite[] = [];
	let rows: Map<string, FormattingRowContext> | undefined;
	if (edit.type === 'format-text') {
		const plan = textFormattingWrites(shape, document, edit, check);
		writes = plan.writes;
		rows = plan.rows;
	} else {
		const plan = await shapeFormattingPlan(pkg, document, shape, edit, check);
		writes = plan.writes;
		rows = plan.rows;
	}
	const changed = new Map<string, Element | undefined>();
	for (const write of writes) {
		check();
		const [section, index, name] = write.name.split('.');
		const rowContext = rows?.get(section!);
		const effective = effectiveShapeCell(shape, document, write.name, write.category, rowContext);
		assertEditableFormattingCell(effective);
		if (
			/^(LinePattern|FillPattern|FillGradientEnabled|LineGradientEnabled|LineColorTrans|FillForegndTrans|FillBkgndTrans|FillBkgnd|ShdwPattern|ShdwForegndTrans|ShapeShdwType|GlowColorTrans|ReflectionTrans|ReflectionSize|QuickStyle[A-Za-z]+)$/.test(
				write.name,
			) &&
			effective?.hasAttribute('U') &&
			visioFormulaCachedValue('0', attribute(effective, 'U')).unit !== 'scalar'
		)
			fail(
				'EDIT_FORMULA_UNIT',
				'Paint patterns, transparency and style selectors require scalar units.',
			);
		if (
			write.unit &&
			!/\.SpLine$/.test(write.name) &&
			effective?.hasAttribute('U') &&
			visioFormulaCachedValue('0', attribute(effective, 'U')).unit !==
				visioFormulaCachedValue('0', write.unit).unit
		)
			fail(
				'EDIT_FORMULA_UNIT',
				'Font size, paragraph indent, line weight and text block cells require matching units.',
			);
		const parent = name ? rowContext!.local.get(index!) : shape;
		const local = parent ? formattingCell(parent, name ?? write.name) : undefined;
		if (
			local &&
			attribute(local, 'F') !== 'Inh' &&
			attribute(local, 'V') === write.value &&
			(!write.formula || attribute(local, 'F') === write.formula) &&
			(!write.dropUnit || !local.hasAttribute('U'))
		)
			continue;
		changed.set(write.name, local);
	}
	if (!changed.size) return false;
	if (edit.type === 'format-shape')
		await assertShapeFormattingPaintScope(pkg, document, shape, edit, changed, check);
	await assertFormattingDependencies(pkg, pagePaths, roots, edit.pageId, shape, changed, check);
	for (const write of writes) {
		check();
		if (!changed.has(write.name)) continue;
		const [sectionName, index, cellName] = write.name.split('.');
		let parent = shape;
		if (cellName) {
			const localRows = rows!.get(sectionName!)!.local;
			let row = localRows.get(index!);
			if (!row) {
				let section = children(shape, 'Section').find(
					(node) => attribute(node, 'N') === sectionName,
				);
				if (!section) {
					section = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Section');
					section.setAttribute('N', sectionName!);
					shape.insertBefore(section, children(shape, 'Text')[0] ?? null);
				}
				row = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Row');
				row.setAttribute('IX', index!);
				section.appendChild(row);
				localRows.set(index!, row);
			}
			parent = row;
		}
		let cell = formattingCell(parent, cellName ?? write.name);
		if (!cell) {
			cell = shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Cell');
			cell.setAttribute('N', cellName ?? write.name);
			parent.insertBefore(
				cell,
				Array.from(parent.childNodes).find(
					(node) => node.nodeType === 1 && (node as Element).localName !== 'Cell',
				) ?? null,
			);
		}
		cell.setAttribute('V', write.value);
		if (write.unit) cell.setAttribute('U', write.unit);
		else if (write.dropUnit) cell.removeAttribute('U');
		if (write.formula) cell.setAttribute('F', write.formula);
		else if (attribute(cell, 'F') !== 'No Formula') cell.removeAttribute('F');
	}
	return true;
}
