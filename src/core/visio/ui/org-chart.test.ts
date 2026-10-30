import { describe, expect, it } from 'vitest';
import {
	createVsdx,
	editVsdx,
	parseVsdx,
	type VisioDocument,
	type VisioEdit,
	type VisioPage,
} from '../index';
import {
	VISIO_ORG_ROWS,
	visioIsOrgAssistant,
	visioIsOrgPosition,
	visioOrgFollowData,
	visioOrgHierarchy,
	visioOrgText,
} from './org-chart';
import { visioOrgAddPosition, visioOrgDeletePosition } from './org-chart-edit';
import { visioOrgImportEdits, visioOrgPeople, type VisioOrgPerson } from './org-chart-import';
import { planOrgLayout, visioOrgReLayout } from './org-chart-layout';
import { createTemplateVsdx } from './templates';

const docOf = (bytes: Uint8Array): Promise<VisioDocument> => parseVsdx(bytes);
const pageOf = async (bytes: Uint8Array): Promise<VisioPage> => (await docOf(bytes)).pages[0]!;
const apply = async (bytes: Uint8Array, edits: VisioEdit[]) => (await editVsdx(bytes, edits)).bytes;
const box = (page: VisioPage, id: string) => {
	const shape = page.shapes.find((item) => item.id === id)!;
	return {
		x: shape.transform[4] + shape.width / 2,
		y: shape.transform[5] + shape.height / 2,
		width: shape.width,
		height: shape.height,
	};
};
const byName = (page: VisioPage, name: RegExp) => page.shapes.filter((s) => name.test(s.name));
/** IDs of the positions in template order: the group of each Belt master takes sixteen IDs. */
const T = (index: number) => String(1 + index * 16);

describe('org chart', () => {
	it('writes the template as instances of Visio org chart masters with their Shape Data', async () => {
		const doc = await docOf(await createTemplateVsdx('organization-chart'));
		const page = doc.pages[0]!;
		expect(doc.masters?.map((m) => m.name)).toEqual([
			'Executive Belt',
			'Manager Belt',
			'Position Belt',
			'Dynamic connector',
		]);
		const hierarchy = visioOrgHierarchy(page);
		// Instances are named after their master, then "Name.ID" as Visio names a later one.
		expect(hierarchy.positions.map((shape) => shape.name.replace(/\.\d+$/, ''))).toEqual([
			'Executive Belt',
			'Manager Belt',
			'Manager Belt',
			'Manager Belt',
			'Position Belt',
			'Position Belt',
		]);
		const top = page.shapes.find((shape) => shape.id === T(0))!;
		expect(top.kind).toBe('group');
		// The text of a position is drawn in front of its photo and belt, as in Visio.
		expect(top.groupDisplayMode).toBe(2);
		expect(top.shapeData?.map((item) => item.name)).toEqual(VISIO_ORG_ROWS.map((r) => r.row));
		expect(top.shapeData?.find((item) => item.name === 'Email')?.label).toBe('E-Mail');
		expect(top.shapeData?.find((item) => item.name === 'Title')?.rawValue).toBe('Chief Executive');
		expect(top.text.plainText).toBe('Name\nChief Executive');
		expect(hierarchy.childrenOf(T(0))).toEqual([T(1), T(2), T(3)]);
		expect(hierarchy.parentOf(T(4))).toBe(T(1));
		expect(
			visioIsOrgPosition(page.shapes.find((shape) => shape.name === 'Dynamic connector')),
		).toBe(false);
		expect(byName(page, /^Dynamic connector/)).toHaveLength(5);
	});

	it('keeps the masters identity so Visio recognises the positions', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const { default: JSZip } = await import('jszip');
		const zip = await JSZip.loadAsync(bytes);
		const masters = await zip.file('visio/masters/masters.xml')!.async('string');
		expect(masters).toContain('UniqueID="{0472F5AB-0002-0000-8E40-00608CF305B2}"');
		expect(masters).toContain('UniqueID="{0472F5AB-0004-0000-8E40-00608CF305B2}"');
		expect(masters).toContain('UniqueID="{0472F5AB-0005-0000-8E40-00608CF305B2}"');
		expect(masters).toContain('BaseID="{EB8A4092-0882-4F2F-A940-FCCE8A34134A}"');
		const media = Object.keys(zip.files).filter((name) => name.startsWith('visio/media/'));
		expect(media).toHaveLength(1);
		const rels = Object.keys(zip.files).filter((name) => /masters\/_rels\/master\d+/.test(name));
		expect(rels.length).toBeGreaterThan(0);
	});

	it('adds a subordinate with data and a glued connector in one batch', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const plan = visioOrgAddPosition(await pageOf(bytes), T(2), 'subordinate')!;
		const page = await pageOf(await apply(bytes, plan.edits));
		const hierarchy = visioOrgHierarchy(page);
		expect(hierarchy.parentOf(plan.newId)).toBe(T(2));
		const added = page.shapes.find((shape) => shape.id === plan.newId)!;
		expect(added.name).toMatch(/^Position Belt/);
		expect(added.shapeData?.find((item) => item.name === 'Title')?.rawValue).toBe('Position');
		expect(added.text.plainText).toBe('Name\nPosition');
		expect(box(page, plan.newId).y).toBeLessThan(box(page, T(2)).y);
		const ids = hierarchy.positions.map((shape) => shape.id);
		for (const a of ids)
			for (const b of ids)
				if (a < b) {
					const [p, q] = [box(page, a), box(page, b)];
					expect(Math.abs(p.x - q.x) > 1.6 || Math.abs(p.y - q.y) > 0.85).toBe(true);
				}
	});

	it('adds a coworker under the same manager and refuses one for the top position', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const before = await pageOf(bytes);
		expect(visioOrgAddPosition(before, T(0), 'peer')).toBeUndefined();
		const plan = visioOrgAddPosition(before, T(4), 'peer')!;
		const page = await pageOf(await apply(bytes, plan.edits));
		expect(visioOrgHierarchy(page).parentOf(plan.newId)).toBe(T(1));
	});

	it('places an assistant beside its manager line, between the manager and its subordinates', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const plan = visioOrgAddPosition(await pageOf(bytes), T(1), 'assistant')!;
		const page = await pageOf(await apply(bytes, plan.edits));
		const assistant = page.shapes.find((shape) => shape.id === plan.newId)!;
		expect(visioIsOrgAssistant(assistant)).toBe(true);
		expect(assistant.name).toMatch(/^Assistant Belt/);
		expect(assistant.shapeData?.find((item) => item.name === 'Title')?.rawValue).toBe('Assistant');
		expect(visioOrgHierarchy(page).parentOf(plan.newId)).toBe(T(1));
		const manager = box(page, T(1));
		const at = box(page, plan.newId);
		// To the right of the line that drops from the manager, not under the manager.
		expect(at.x - at.width / 2).toBeGreaterThan(manager.x);
		// Between the manager and its subordinates.
		for (const sub of [T(4), T(5)]) {
			expect(at.y).toBeLessThan(manager.y - manager.height / 2);
			expect(at.y).toBeGreaterThan(box(page, sub).y + box(page, sub).height / 2);
		}
	});

	it('deletes a position with its connectors and reattaches its subordinates in one batch', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const edits = visioOrgDeletePosition(await pageOf(bytes), T(1))!;
		const page = await pageOf(await apply(bytes, edits));
		const hierarchy = visioOrgHierarchy(page);
		expect(page.shapes.some((shape) => shape.id === T(1))).toBe(false);
		expect(hierarchy.positions).toHaveLength(5);
		expect(hierarchy.parentOf(T(4))).toBe(T(0));
		expect(hierarchy.parentOf(T(5))).toBe(T(0));
		expect(page.connectors.every((c) => page.shapes.some((s) => s.id === c.toShapeId))).toBe(true);
	});

	it('lays a scattered tree out again', async () => {
		const bytes = await createTemplateVsdx('organization-chart');
		const moved = await apply(bytes, [
			{ type: 'move-shape', pageId: '0', shapeId: T(4), x: 9, y: 1 },
		]);
		const plan = visioOrgReLayout(await pageOf(moved));
		expect(plan?.length).toBeGreaterThan(0);
		const page = await pageOf(await apply(moved, plan!));
		expect(box(page, T(4)).y).toBeLessThan(box(page, T(1)).y);
		expect(Math.abs(box(page, T(4)).x - 9)).toBeGreaterThan(1);
	});

	it('shows the Name and Title of a position as its text and follows changes to them', async () => {
		expect(visioOrgText({ Name: 'Ada Lovelace', Title: 'Engineer' })).toBe(
			'Ada Lovelace\nEngineer',
		);
		expect(visioOrgText({ Name: '', Title: 'Engineer' })).toBe('Engineer');
		const bytes = await createTemplateVsdx('organization-chart');
		const page = await pageOf(bytes);
		const set = (row: string, value: string, shapeId = T(4)): VisioEdit => ({
			type: 'set-shape-data',
			pageId: '0',
			shapeId,
			row,
			data: { label: row, prompt: '', type: 'string', format: '', value },
		});
		const edits = visioOrgFollowData(page, [set('Name', 'Ada Lovelace')]);
		expect(edits.at(-1)).toMatchObject({
			type: 'replace-plain-text',
			shapeId: T(4),
			text: 'Ada Lovelace\nPosition',
		});
		const after = await pageOf(await apply(bytes, edits));
		const shape = after.shapes.find((item) => item.id === T(4))!;
		expect(shape.text.plainText).toBe('Ada Lovelace\nPosition');
		expect(shape.shapeData?.find((item) => item.name === 'Name')?.rawValue).toBe('Ada Lovelace');
		// Rows that are not shown, and shapes that are not positions, add no text edit.
		expect(visioOrgFollowData(page, [set('Telephone', '555')])).toHaveLength(1);
		expect(visioOrgFollowData(page, [set('Name', 'x', '999')])).toHaveLength(1);
	});

	it('builds a chart from Name, Title and Reports To rows in one batch', async () => {
		const table = {
			columns: ['Name', 'Title', 'Reports To', 'Department', 'Type'].map((name) => ({
				name,
				label: name,
				type: 'string' as const,
			})),
			rows: [
				['Ada', 'CEO', '', 'Board', ''],
				['Bo', 'CTO', 'Ada', 'Tech', ''],
				['Cy', 'CFO', 'Ada', 'Finance', ''],
				['Di', 'Engineer', 'Bo', 'Tech', ''],
				['Ed', 'Assistant', 'Ada', '', 'Assistant'],
			],
		};
		const people = visioOrgPeople(table) as VisioOrgPerson[];
		expect(people).toHaveLength(5);
		expect(people[4]).toMatchObject({ name: 'Ed', assistant: true });
		const blank = await createVsdx({ width: 11, height: 8.5 });
		const plan = visioOrgImportEdits(await pageOf(blank), people);
		if (typeof plan === 'string') throw new Error(plan);
		const page = await pageOf(await apply(blank, plan.edits));
		const hierarchy = visioOrgHierarchy(page);
		expect(hierarchy.positions).toHaveLength(5);
		const named = (name: string) =>
			hierarchy.positions.find(
				(shape) => shape.shapeData?.find((row) => row.name === 'Name')?.rawValue === name,
			)!;
		expect(named('Ada').name).toMatch(/^Executive Belt/);
		expect(named('Bo').name).toMatch(/^Manager Belt/);
		expect(named('Di').name).toMatch(/^Position Belt/);
		expect(named('Ed').name).toMatch(/^Assistant Belt/);
		expect(hierarchy.parentOf(named('Di').id)).toBe(named('Bo').id);
		expect(hierarchy.parentOf(named('Ed').id)).toBe(named('Ada').id);
		expect(named('Bo').text.plainText).toBe('Bo\nCTO');
		expect(named('Bo').shapeData?.find((row) => row.name === 'Department')?.rawValue).toBe('Tech');
		expect(box(page, named('Ada').id).y).toBeGreaterThan(box(page, named('Bo').id).y);
	});

	it('reads people from a sheet by heading and reports what is missing', () => {
		const column = (name: string) => ({ name, label: name, type: 'string' as const });
		expect(visioOrgPeople({ columns: [column('Name'), column('Title')], rows: [['a', 'b']] })).toBe(
			'The sheet needs a Reports To column with a heading in its first row.',
		);
		const people = visioOrgPeople({
			columns: ['Employee ID', 'Full Name', 'Job Title', 'Manager', 'E-Mail'].map(column),
			rows: [
				['1', 'Ada', 'CEO', '', 'ada@example.com'],
				['2', 'Bo', 'CTO', '1', ''],
				['', '', '', '', ''],
			],
		}) as VisioOrgPerson[];
		expect(people).toHaveLength(2);
		expect(people[0]).toMatchObject({ id: '1', email: 'ada@example.com' });
		const none = visioOrgPeople({
			columns: ['Name', 'Title', 'Reports To'].map(column),
			rows: [],
		});
		expect(none).toBe('The sheet has no people.');
	});

	it('imports a manager cycle without hanging and keeps every person', async () => {
		const blank = await createVsdx({ width: 11, height: 8.5 });
		const people: VisioOrgPerson[] = [
			{ name: 'A', title: 't', reportsTo: 'B' },
			{ name: 'B', title: 't', reportsTo: 'A' },
		];
		const plan = visioOrgImportEdits(await pageOf(blank), people);
		if (typeof plan === 'string') throw new Error(plan);
		const page = await pageOf(await apply(blank, plan.edits));
		expect(visioOrgHierarchy(page).positions).toHaveLength(2);
	});

	it('lays a tree out with each manager centred over its subordinates and assistants in a band', () => {
		const node = (id: string, assistant = false) => ({
			id,
			width: 1.6,
			height: 0.9,
			...(assistant ? { assistant } : {}),
		});
		const nodes = [node('a'), node('b'), node('c'), node('d'), node('e'), node('x', true)];
		const edges = [
			['a', 'b'],
			['a', 'c'],
			['b', 'd'],
			['b', 'e'],
			['b', 'x'],
		].map(([from, to]) => ({ from: from!, to: to! }));
		const at = planOrgLayout(nodes, edges)!;
		const get = (id: string) => at.get(id)!;
		expect(get('a').x).toBeCloseTo((get('b').x + get('c').x) / 2);
		expect(get('b').x).toBeCloseTo((get('d').x + get('e').x) / 2);
		// The assistant is right of its manager's line and in a band between the two rows.
		expect(get('x').x).toBeGreaterThan(get('b').x + 0.8);
		expect(get('x').y).toBeGreaterThan(get('b').y + 0.9);
		expect(get('x').y).toBeLessThan(get('d').y - 0.9);
		// Every level is a row, and nothing overlaps.
		expect(get('b').y).toBeCloseTo(get('c').y);
		expect(get('d').y).toBeCloseTo(get('e').y);
		for (const p of nodes)
			for (const q of nodes)
				if (p.id < q.id)
					expect(
						Math.abs(get(p.id).x - get(q.id).x) >= 1.6 ||
							Math.abs(get(p.id).y - get(q.id).y) >= 0.9,
					).toBe(true);
		// Positions with no tops, or in a cycle, are still placed.
		expect(
			planOrgLayout(
				[node('p'), node('q')],
				[
					{ from: 'p', to: 'q' },
					{ from: 'q', to: 'p' },
				],
			)?.size,
		).toBe(2);
		expect(planOrgLayout([], [])).toBeUndefined();
	});

	it('grows the page for a chart that does not fit and places a second import beside it', async () => {
		const blank = await createVsdx({ width: 11, height: 8.5 });
		const people: VisioOrgPerson[] = Array.from({ length: 12 }, (_, index) => ({
			name: `P${index}`,
			title: 'T',
			reportsTo: index ? `P${Math.floor((index - 1) / 3)}` : '',
		}));
		const first = visioOrgImportEdits(await pageOf(blank), people);
		if (typeof first === 'string') throw new Error(first);
		const bytes = await apply(blank, first.edits);
		const page = await pageOf(bytes);
		expect(page.width).toBeGreaterThan(11);
		for (const shape of visioOrgHierarchy(page).positions)
			expect(box(page, shape.id).x - shape.width / 2).toBeGreaterThanOrEqual(0);
		const right = Math.max(...page.shapes.map((shape) => box(page, shape.id).x + shape.width / 2));
		const second = visioOrgImportEdits(page, [{ name: 'Z', title: 'T', reportsTo: '' }]);
		if (typeof second === 'string') throw new Error(second);
		const both = await pageOf(await apply(bytes, second.edits));
		const added = both.shapes.find((shape) => shape.id === second.firstId)!;
		expect(box(both, added.id).x - added.width / 2).toBeGreaterThan(right);
		// Too many people for one batch say so instead of failing half way.
		const many = Array.from({ length: 400 }, (_, index) => ({
			name: `N${index}`,
			title: 'T',
			reportsTo: '',
		}));
		expect(visioOrgImportEdits(page, many)).toMatch(/import at most about/);
	});
});
