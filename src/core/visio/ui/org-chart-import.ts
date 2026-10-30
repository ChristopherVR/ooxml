import type { VisioEdit } from '../edit-commands';
import type { VisioPage } from '../model';
import type { VisioDataTable } from './data-table';
import { visioShapePageBounds } from './diagram-parts';
import {
	VISIO_ORG_SIZE,
	visioOrgDropPosition,
	visioOrgFreeIds,
	visioOrgReportingLine,
	type VisioOrgMaster,
} from './org-chart-edit';
import { planOrgLayout } from './org-chart-layout';
import { visioOrgDataEdits, visioOrgText } from './org-chart';

/** One row of an org chart import: a person and who they report to. */
export interface VisioOrgPerson {
	name: string;
	title: string;
	/** The Name (or ID) of the manager; empty for the top of a tree. */
	reportsTo: string;
	/** An identifier other rows may use in their Reports To column. */
	id?: string;
	department?: string;
	telephone?: string;
	email?: string;
	assistant?: boolean;
}

/** Most edits one import may use (the editor's per-call limit is 1000). */
const MAX_EDITS = 980;
/** Page inches kept clear around an imported chart. */
const MARGIN = 0.5;

const COLUMN_NAMES = {
	name: ['name', 'employee', 'employee name', 'full name', 'person'],
	title: ['title', 'job title', 'position', 'role'],
	reportsTo: ['reports to', 'reports_to', 'reportsto', 'manager', 'supervisor', 'boss'],
	id: ['id', 'employee id', 'employee number', 'unique id', 'unique identifier', 'emp id'],
	department: ['department', 'dept'],
	telephone: ['telephone', 'phone', 'phone number', 'tel'],
	email: ['email', 'e-mail', 'e mail', 'email address'],
	type: ['type', 'shape type', 'kind'],
} as const;

/**
 * The people of an imported sheet. Columns are found by their headings: Name, Title and Reports To
 * are required (Reports To holds the manager's Name, or their ID when the sheet has an ID column);
 * Department, Telephone, E-Mail and a Type column (a value of "Assistant" makes the row an
 * assistant) are used when present. Returns an error sentence when the sheet cannot be read.
 */
export function visioOrgPeople(table: VisioDataTable): VisioOrgPerson[] | string {
	const find = (names: readonly string[]) =>
		table.columns.findIndex((column) => names.includes(column.name.trim().toLowerCase()));
	const at = Object.fromEntries(
		(Object.keys(COLUMN_NAMES) as (keyof typeof COLUMN_NAMES)[]).map((key) => [
			key,
			find(COLUMN_NAMES[key]),
		]),
	) as Record<keyof typeof COLUMN_NAMES, number>;
	const missing = (['name', 'title', 'reportsTo'] as const).filter((key) => at[key] < 0);
	if (missing.length)
		return `The sheet needs ${missing
			.map((key) => ({ name: 'a Name', title: 'a Title', reportsTo: 'a Reports To' })[key])
			.join(', ')} column with a heading in its first row.`;
	const cell = (row: string[], index: number) => (index < 0 ? '' : (row[index] ?? '').trim());
	const people = table.rows
		.filter((row) => cell(row, at.name))
		.map((row): VisioOrgPerson => {
			const person: VisioOrgPerson = {
				name: cell(row, at.name),
				title: cell(row, at.title),
				reportsTo: cell(row, at.reportsTo),
			};
			for (const [key, index] of [
				['id', at.id],
				['department', at.department],
				['telephone', at.telephone],
				['email', at.email],
			] as const)
				if (cell(row, index)) person[key] = cell(row, index);
			if (/^assistant$/i.test(cell(row, at.type))) person.assistant = true;
			return person;
		});
	return people.length ? people : 'The sheet has no people.';
}

interface Node {
	person: VisioOrgPerson;
	index: number;
	parent?: number;
	children: number[];
}
/** The reporting forest: each person's manager is found by ID, then by name; cycles are broken. */
function forest(people: readonly VisioOrgPerson[]): Node[] {
	const nodes: Node[] = people.map((person, index) => ({ person, index, children: [] }));
	const lookup = (key: string): number | undefined => {
		const wanted = key.trim().toLowerCase();
		if (!wanted) return undefined;
		const found =
			nodes.find((node) => node.person.id?.toLowerCase() === wanted) ??
			nodes.find((node) => node.person.name.toLowerCase() === wanted);
		return found?.index;
	};
	for (const node of nodes) {
		const parent = lookup(node.person.reportsTo);
		if (parent === undefined || parent === node.index) continue;
		// A person who ends up their own manager through a chain stays the top of their tree.
		let up: number | undefined = parent;
		let cyclic = false;
		for (let steps = 0; up !== undefined && steps <= nodes.length; steps++) {
			if (up === node.index) {
				cyclic = true;
				break;
			}
			up = nodes[up]!.parent;
		}
		if (cyclic) continue;
		node.parent = parent;
		nodes[parent]!.children.push(node.index);
	}
	return nodes;
}

/**
 * Org Chart > Import: build an org chart from `people` on `page`, to the right of what is already
 * there. Tops of trees are Executive Belts, people with subordinates Manager Belts, the rest
 * Position Belts and assistants Assistant Belts, joined by Dynamic connectors, with Name, Title
 * and the optional rows as Shape Data and the text following the data. The page grows to hold
 * the chart. One batch, so it is one undo step. Returns an error sentence when it cannot be done.
 */
export function visioOrgImportEdits(
	page: VisioPage,
	people: readonly VisioOrgPerson[],
): { edits: VisioEdit[]; firstId: string } | string {
	const nodes = forest(people);
	const lines = nodes.filter((node) => node.parent !== undefined);
	const count = nodes.length;
	const perPerson = (node: Node) => 2 + visioOrgDataEdits('', '', dataOf(node.person)).length;
	const total = nodes.reduce((sum, node) => sum + perPerson(node), 0) + lines.length * 3;
	if (total > MAX_EDITS)
		return `That chart needs ${total} edits; import at most about ${Math.floor(
			(count * MAX_EDITS) / total,
		)} people at a time.`;
	const ids = visioOrgFreeIds(page, count, lines.length);
	if (!ids) return 'The page has no free shape IDs.';
	const ratio = page.drawingToPageScale ?? 1;
	const size = { width: VISIO_ORG_SIZE.width * ratio, height: VISIO_ORG_SIZE.height * ratio };
	const placed = planOrgLayout(
		nodes.map((node) => ({
			id: String(node.index),
			...size,
			...(node.person.assistant && node.parent !== undefined ? { assistant: true } : {}),
		})),
		lines.map((node) => ({ from: String(node.parent), to: String(node.index) })),
	);
	if (!placed) return 'The chart could not be laid out.';
	const centres = [...placed.values()];
	const chartWidth = Math.max(...centres.map((at) => at.x)) + size.width / 2;
	const chartHeight = Math.max(...centres.map((at) => at.y)) + size.height / 2;
	const chartLeft = Math.min(...centres.map((at) => at.x)) - size.width / 2;
	let right = 0;
	for (const shape of page.shapes) right = Math.max(right, visioShapePageBounds(shape).maxX);
	const left = (right > 0 ? right + MARGIN : MARGIN) - chartLeft;
	// A page that holds glued connectors cannot be resized by the editor yet, so a chart that does
	// not fit overhangs it there.
	const grow = page.connectors.length === 0;
	const pageWidth = grow ? Math.max(page.width, left + chartWidth + MARGIN) : page.width;
	const pageHeight = grow ? Math.max(page.height, chartHeight + 2 * MARGIN) : page.height;
	const pageId = page.id;
	const centre = (node: Node) => {
		const at = placed.get(String(node.index))!;
		return { x: (left + at.x) / ratio, y: (pageHeight - MARGIN - at.y) / ratio };
	};
	const master = (node: Node): VisioOrgMaster =>
		node.person.assistant && node.parent !== undefined
			? 'Assistant Belt'
			: node.parent === undefined
				? 'Executive Belt'
				: node.children.length
					? 'Manager Belt'
					: 'Position Belt';
	const drops = nodes.map((node) =>
		visioOrgDropPosition(pageId, ids.positions[node.index]!, master(node), centre(node)),
	);
	const reporting = lines.map((node, index) => {
		const a = centre(nodes[node.parent!]!);
		const b = centre(node);
		return visioOrgReportingLine(
			pageId,
			ids.lines[index]!,
			ids.positions[node.parent!]!,
			ids.positions[node.index]!,
			{ x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 },
		);
	});
	const edits: VisioEdit[] = [
		...(pageWidth > page.width || pageHeight > page.height
			? [{ type: 'set-page-size' as const, pageId, width: pageWidth, height: pageHeight }]
			: []),
		...drops,
		...reporting.map((line) => line.drop),
		...nodes.flatMap((node) => [
			{
				type: 'replace-plain-text' as const,
				pageId,
				shapeId: ids.positions[node.index]!,
				text: visioOrgText({ Name: node.person.name, Title: node.person.title }),
			},
			...visioOrgDataEdits(pageId, ids.positions[node.index]!, dataOf(node.person)),
		]),
		...reporting.flatMap((line) => line.glue),
	];
	return { edits, firstId: ids.positions[0]! };
}

const dataOf = (person: VisioOrgPerson) => ({
	Name: person.name,
	Title: person.title,
	...(person.department ? { Department: person.department } : {}),
	...(person.telephone ? { Telephone: person.telephone } : {}),
	...(person.email ? { Email: person.email } : {}),
});
