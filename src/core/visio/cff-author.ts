import { NS, parseXml } from '../xml/index';
import type { VisioPackageLimits } from './package-common';
import { setCell } from './edit-geometry-cells';
import { addMasterToDrawing } from './edit-master-add';
import { userSection } from './edit-diagram-parts';
import { applyShapeCreationStyles } from './edit-shape-create';
import { CFF_AUTHORED_IDS, type CffRole } from './cff-masters';

type Row = [string, string, string?];

interface CffMasterDefinition {
	name: string;
	size: { width: number; height: number };
	/** The rows that make a shape this kind of part: the categories and structure Visio checks. */
	rows: Row[];
}

const MM = 1 / 25.4;
const STENCIL: Row = ['visMasterStencil', 'xfunc_m.vss'];
const VERSION: Row = ['visVersion', '15', ''];
const HIDDEN_LAYER: Row = ['msvSDMembersOnHiddenLayer', '1', 'BOOL'];
const QUIET: Row[] = [
	['msvSDContainerNoHighlight', '1', 'BOOL'],
	['msvSDContainerNoRibbon', '1', 'BOOL'],
];
const ALL_PARTS = 'CFF Container;CFF List;Swimlane;Phase;NonCFF';

/**
 * Visio's five hidden masters of a cross-functional flowchart, by name, with the User rows that
 * identify each part (the values are the ones Visio's masters hold). The geometry, the formulas
 * that tie a master to its container and Visio's heading shapes are not reproduced: every
 * instance this package writes carries its own cells, and these masters tell Visio which part
 * of the chart the instance is.
 */
const DEFINITIONS: Record<CffRole, CffMasterDefinition> = {
	frame: {
		name: 'CFF Container',
		size: { width: 150 * MM, height: 100 * MM },
		rows: [
			['msvShapeCategories', 'CFF Container;DoNotContain'],
			['msvStructureType', 'Container'],
			...QUIET,
			['msvSDContainerRequiredCategories', 'CFF List'],
			['msvSDContainerExcludedCategories', ''],
			STENCIL,
			VERSION,
		],
	},
	'swimlane-list': {
		name: 'Swimlane List',
		size: { width: 125 * MM, height: 100 * MM },
		rows: [
			['msvStructureType', 'List'],
			['msvShapeCategories', 'CFF List;DoNotContain'],
			['msvSDListRequiredCategories', 'Swimlane'],
			['msvSDContainerExcludedCategories', 'CFF Container;CFF List;Phase;NonCFF'],
			...QUIET,
			['msvSDListDirection', '2', ''],
			HIDDEN_LAYER,
			STENCIL,
			VERSION,
		],
	},
	'phase-list': {
		name: 'Phase List',
		size: { width: 100 * MM, height: 150 * MM },
		rows: [
			['msvStructureType', 'List'],
			['msvShapeCategories', 'CFF List;DoNotContain'],
			['msvSDContainerRequiredCategories', 'Phase'],
			['msvSDListRequiredCategories', 'Phase'],
			...QUIET,
			['msvSDListDirection', '0', ''],
			HIDDEN_LAYER,
			STENCIL,
			VERSION,
		],
	},
	lane: {
		name: 'Swimlane',
		size: { width: 125 * MM, height: 30 * MM },
		rows: [
			['msvShapeCategories', 'Swimlane;DoNotContain'],
			['msvStructureType', 'Container'],
			['msvSDContainerNoRibbon', '1', 'BOOL'],
			['msvSDContainerExcludedCategories', ALL_PARTS],
			HIDDEN_LAYER,
			STENCIL,
			VERSION,
		],
	},
	separator: {
		name: 'Separator',
		size: { width: 0.1, height: 100 * MM },
		rows: [
			['msvStructureType', 'Container'],
			['msvShapeCategories', 'Phase;DoNotContain'],
			['msvSDContainerExcludedCategories', ALL_PARTS],
			...QUIET,
			HIDDEN_LAYER,
			STENCIL,
			VERSION,
		],
	},
};

/** The order Visio's template lists them in. */
export const CFF_ROLES: readonly CffRole[] = [
	'frame',
	'swimlane-list',
	'phase-list',
	'lane',
	'separator',
];

/** The base ID is the master's family; ours differs from the unique ID in its first group. */
const baseId = (uniqueId: string) => uniqueId.replace(/^\{[0-9A-F]{8}/, '{4F0C7A10');

function entry(role: CffRole) {
	return (masters: Element, id: string, name: string, relationshipId: string): void => {
		const doc = masters.ownerDocument!;
		const node = (parent: Element, tag: string) =>
			parent.appendChild(doc.createElementNS(masters.namespaceURI, tag)) as Element;
		const master = node(masters, 'Master');
		const { size } = DEFINITIONS[role];
		for (const [attribute, value] of [
			['ID', id],
			['NameU', name],
			['IsCustomNameU', '1'],
			['Name', name],
			['IsCustomName', '1'],
			['Prompt', ''],
			['IconSize', '1'],
			['AlignName', '2'],
			['MatchByName', '0'],
			['IconUpdate', '1'],
			['UniqueID', CFF_AUTHORED_IDS[role]],
			['BaseID', baseId(CFF_AUTHORED_IDS[role])],
			['PatternFlags', '0'],
			['Hidden', '1'],
			['MasterType', '34'],
		] as const)
			master.setAttribute(attribute, value);
		const sheet = node(master, 'PageSheet');
		for (const style of ['LineStyle', 'FillStyle', 'TextStyle']) sheet.setAttribute(style, '0');
		for (const [cell, value] of [
			['PageWidth', size.width],
			['PageHeight', size.height],
			['PageScale', 1],
			['DrawingScale', 1],
			['DrawingSizeType', 4],
			['DrawingScaleType', 0],
		] as const)
			setCell(sheet, cell, value);
		node(master, 'Rel').setAttributeNS(NS.r, 'r:id', relationshipId);
	};
}

function contents(role: CffRole) {
	return (document: Element): Element => {
		const namespace = document.namespaceURI!;
		const root = parseXml(`<MasterContents xmlns="${namespace}"/>`).documentElement;
		root.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
		const doc = root.ownerDocument!;
		const shapes = root.appendChild(doc.createElementNS(namespace, 'Shapes')) as Element;
		const shape = shapes.appendChild(doc.createElementNS(namespace, 'Shape')) as Element;
		shape.setAttribute('ID', '5');
		shape.setAttribute('Type', 'Shape');
		applyShapeCreationStyles(shape, document);
		const { width, height } = DEFINITIONS[role].size;
		for (const [name, value, formula] of [
			['PinX', width / 2],
			['PinY', height / 2],
			['Width', width],
			['Height', height],
			['LocPinX', width / 2, 'Width*0.5'],
			['LocPinY', height / 2, 'Height*0.5'],
			['Angle', 0],
			['FlipX', 0],
			['FlipY', 0],
			['ResizeMode', 0],
		] as const)
			setCell(shape, name, value, formula);
		shape.appendChild(userSection(shape, DEFINITIONS[role].rows));
		return root;
	};
}

/**
 * Add Visio's hidden cross-functional flowchart masters (CFF Container, Swimlane List, Phase
 * List, Swimlane, Separator) to the drawing, as Visio's template does, so the shapes of a chart
 * this package creates can be instances of them. Masters added earlier are found by UniqueID.
 */
export async function ensureCffMasters(
	source: Uint8Array,
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
	check: () => void,
): Promise<{ bytes: Uint8Array; changedParts: string[] }> {
	let bytes = source;
	const changedParts = new Set<string>();
	for (const role of CFF_ROLES) {
		const added = await addMasterToDrawing(
			bytes,
			{
				uniqueId: CFF_AUTHORED_IDS[role],
				name: DEFINITIONS[role].name,
				entry: entry(role),
				contents: contents(role),
			},
			limits,
			maxOutput,
			deadline,
			check,
		);
		bytes = added.bytes;
		for (const part of added.changedParts) changedParts.add(part);
	}
	return { bytes, changedParts: [...changedParts] };
}
