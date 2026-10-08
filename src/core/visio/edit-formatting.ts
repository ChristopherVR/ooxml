import type { VisioPackage } from './package';
import type { VisioFormatEdit } from './edit-formatting-commands';
import { attribute, children } from './sheet';
import { fail } from './package-common';
import { visioFormulaCachedValue } from './formula';
import { assertFormattingDependencies } from './edit-formatting-scope';
import {
	assertEditableFormattingCell,
	assertShapeLocks,
	effectiveShapeCell,
	formattingRow,
	formattingCell,
	type FormattingCategory,
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
function plainUniformText(shape: Element): void {
	const texts = children(shape, 'Text');
	if (
		texts.length !== 1 ||
		children(shape, 'Section').some((node) => attribute(node, 'N') === 'Field')
	)
		fail(
			'UNSUPPORTED_FORMAT_EDIT',
			'Text formatting requires one existing local Text without fields.',
		);
	for (const part of Array.from(texts[0]!.childNodes)) {
		if (part.nodeType === 3 || part.nodeType === 4) continue;
		const marker = part as Element;
		if (
			part.nodeType !== 1 ||
			marker.namespaceURI !== shape.namespaceURI ||
			!['cp', 'pp', 'tp'].includes(marker.localName) ||
			(attribute(marker, 'IX') ?? '0') !== '0' ||
			marker.childNodes.length
		)
			fail(
				'UNSUPPORTED_FORMAT_EDIT',
				'Rich text, nonzero markers and unknown text markup cannot be formatted.',
			);
	}
	formattingRow(shape, 'Character');
	formattingRow(shape, 'Paragraph');
}
interface Write {
	name: string;
	value: string;
	unit?: string;
	formula?: string;
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
	for (const category of ['LineStyle', 'FillStyle', 'TextStyle'] as const)
		if (effectiveShapeCell(shape, document, 'LayerMember', category))
			fail('UNSUPPORTED_FORMAT_EDIT', 'Layered shape formatting is not yet supported.');
	assertShapeLocks(
		shape,
		document,
		edit.type === 'format-text' ? ['LockFormat', 'LockTextEdit'] : ['LockFormat'],
	);
	const writes: Write[] = [];
	const add = (
		name: string,
		value: string | number,
		category: FormattingCategory,
		unit?: string,
		formula?: string,
	) =>
		writes.push({
			name,
			value: String(value),
			category,
			...(unit ? { unit } : {}),
			...(formula ? { formula } : {}),
		});
	if (edit.type === 'format-text') {
		plainUniformText(shape);
		if (edit.fontSize !== undefined) add('Character.0.Size', edit.fontSize / 72, 'TextStyle', 'PT');
		if (edit.fontFamily !== undefined) {
			const names = children(document, 'FaceNames');
			if (names.length !== 1)
				fail('UNSUPPORTED_FORMAT_EDIT', 'Font family requires existing document FaceNames.');
			const fonts = children(names[0], 'FaceName');
			const ids = new Set<string>();
			for (const font of fonts) {
				const id = attribute(font, 'ID');
				if (!id || !/^\d+$/.test(id) || ids.has(id))
					fail('UNSUPPORTED_FORMAT_EDIT', 'Document font IDs must be explicit and unique.');
				ids.add(id);
			}
			const matches = fonts.filter(
				(font) =>
					(attribute(font, 'Name') ?? attribute(font, 'NameU'))?.toLowerCase() ===
					edit.fontFamily!.toLowerCase(),
			);
			if (matches.length !== 1)
				fail('UNSUPPORTED_FORMAT_EDIT', 'Font family must match one existing document FaceName.');
			add('Character.0.Font', attribute(matches[0], 'ID')!, 'TextStyle');
		}
		if ([edit.bold, edit.italic, edit.underline].some((value) => value !== undefined)) {
			const cell = effectiveShapeCell(shape, document, 'Character.0.Style', 'TextStyle');
			const cached = cell
				? visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U'))
				: { value: 0, unit: 'scalar' };
			if (
				cached.unit !== 'scalar' ||
				!Number.isSafeInteger(cached.value) ||
				cached.value < 0 ||
				cached.value > 255
			)
				fail('UNSUPPORTED_FORMAT_EDIT', 'Character style requires a valid scalar bit mask.');
			let bits = cached.value;
			for (const [value, bit] of [
				[edit.bold, 1],
				[edit.italic, 2],
				[edit.underline, 4],
			] as const)
				if (value !== undefined) bits = value ? bits | bit : bits & ~bit;
			add('Character.0.Style', bits, 'TextStyle');
		}
		if (edit.horizontalAlign !== undefined)
			add(
				'Paragraph.0.HorzAlign',
				['left', 'center', 'right'].indexOf(edit.horizontalAlign),
				'TextStyle',
			);
		if (edit.verticalAlign !== undefined)
			add('VerticalAlign', ['top', 'middle', 'bottom'].indexOf(edit.verticalAlign), 'TextStyle');
	} else {
		if (edit.fillColor !== undefined) {
			add('FillPattern', edit.fillColor === 'none' ? 0 : 1, 'FillStyle');
			add('FillGradientEnabled', 0, 'FillStyle');
			if (edit.fillColor !== 'none') {
				add('FillForegnd', edit.fillColor, 'FillStyle', undefined, rgb(edit.fillColor));
				add('FillForegndTrans', 0, 'FillStyle');
			}
		}
		if (edit.lineColor !== undefined) {
			add('LineColor', edit.lineColor, 'LineStyle', undefined, rgb(edit.lineColor));
			add('LineColorTrans', 0, 'LineStyle');
			add('LineGradientEnabled', 0, 'LineStyle');
		}
		if (edit.lineWeight !== undefined) add('LineWeight', edit.lineWeight / 72, 'LineStyle', 'PT');
	}
	const changed = new Map<string, Element | undefined>();
	for (const write of writes) {
		const effective = effectiveShapeCell(shape, document, write.name, write.category);
		assertEditableFormattingCell(effective);
		if (
			write.unit &&
			effective?.hasAttribute('U') &&
			visioFormulaCachedValue('0', attribute(effective, 'U')).unit !== 'length'
		)
			fail('EDIT_FORMULA_UNIT', 'Font size and line weight require length units.');
		const [section, , name] = write.name.split('.');
		const parent = name ? formattingRow(shape, section!) : shape;
		const local = parent ? formattingCell(parent, name ?? write.name) : undefined;
		if (
			local &&
			attribute(local, 'V') === write.value &&
			(!write.formula || attribute(local, 'F') === write.formula)
		)
			continue;
		changed.set(write.name, local);
	}
	if (!changed.size) return false;
	await assertFormattingDependencies(pkg, pagePaths, roots, edit.pageId, shape, changed, check);
	for (const write of writes) {
		if (!changed.has(write.name)) continue;
		const [sectionName, , cellName] = write.name.split('.');
		let parent = shape;
		if (cellName) {
			let row = formattingRow(shape, sectionName!);
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
				row.setAttribute('IX', '0');
				section.appendChild(row);
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
		if (write.formula) cell.setAttribute('F', write.formula);
		else if (attribute(cell, 'F') !== 'No Formula') cell.removeAttribute('F');
	}
	return true;
}
function rgb(color: string): string {
	return `RGB(${[1, 3, 5].map((index) => parseInt(color.slice(index, index + 2), 16)).join(',')})`;
}
