// Applies one edit per supported family to the Visio-authored reference corpus with the core API and
// writes <case>.vsdx plus cases.json (the cell values our written XML claims) for
// scripts/record-visio-edit-families.ps1 to compare against a real Visio instance.
// Run: bun scripts/create-visio-edit-family-fixtures.ts <OutputDirectory>
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { editVsdx } from '../src/core/visio/edit';
import { captureVisioClipboard } from '../src/core/visio/clipboard';
import type { VisioEdit } from '../src/core/visio/edit-commands';

const directory = process.argv[2];
if (!directory) throw new Error('Provide a directory for owned native test files.');
const output = resolve(directory);
await mkdir(output, { recursive: true });
// Second argument: reference-flat.vsdx (default) or reference-corpus.vsdx.
const source = process.argv[3] ?? 'reference-flat.vsdx';
const flatDocument = source === 'reference-flat.vsdx';
const prefix = flatDocument ? 'flat-' : '';
const corpus = new Uint8Array(
	await readFile(
		resolve(import.meta.dir, '../viewers/visio/docs/evidence/reference-corpus', source),
	),
);

const kids = (node: Element, name: string): Element[] =>
	Array.from(node.childNodes as unknown as Element[]).filter((n) => n.nodeName === name);
async function pageMap(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const pages = new DOMParser().parseFromString(
		await zip.file('visio/pages/pages.xml')!.async('string'),
		'text/xml',
	);
	const rels = await zip.file('visio/pages/_rels/pages.xml.rels')!.async('string');
	const result = new Map<string, { id: string; path: string }>();
	for (const page of Array.from(pages.getElementsByTagName('Page'))) {
		const rid = kids(page, 'Rel')[0]!.getAttribute('r:id')!;
		const target = new RegExp(
			`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`,
		).exec(rels)!;
		result.set(page.getAttribute('NameU')!, {
			id: page.getAttribute('ID')!,
			path: 'visio/pages/' + (target[1] ?? target[2]),
		});
	}
	return { zip, result };
}
const NUMERIC = [
	'PinX',
	'PinY',
	'Width',
	'Height',
	'Angle',
	'LocPinX',
	'LocPinY',
	'LineWeight',
	'FillPattern',
	'LinePattern',
	'FlipX',
	'FlipY',
];
/** Cached values our package claims for every shape of every page (nested shapes included). */
async function claims(bytes: Uint8Array) {
	const { zip, result } = await pageMap(bytes);
	const pages: Record<string, unknown>[] = [];
	for (const [name, info] of result) {
		const doc = new DOMParser().parseFromString(
			await zip.file(info.path)!.async('string'),
			'text/xml',
		);
		const shapes: Record<string, unknown>[] = [];
		for (const shape of Array.from(doc.getElementsByTagName('Shape'))) {
			const values: Record<string, number> = {};
			const colors: Record<string, string> = {};
			for (const cell of kids(shape, 'Cell')) {
				const n = cell.getAttribute('N')!;
				const v = Number(cell.getAttribute('V'));
				if (NUMERIC.includes(n) && cell.hasAttribute('V') && Number.isFinite(v)) values[n] = v;
				const hex = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(
					cell.getAttribute('V') ?? '',
				);
				if ((n === 'FillForegnd' || n === 'LineColor') && hex)
					colors[n] =
						`RGB(${parseInt(hex[1]!, 16)},${parseInt(hex[2]!, 16)},${parseInt(hex[3]!, 16)})`;
			}
			for (const section of kids(shape, 'Section').filter(
				(s) => s.getAttribute('N') === 'Character',
			)) {
				const row = kids(section, 'Row')[0];
				const size = row && kids(row, 'Cell').find((c) => c.getAttribute('N') === 'Size');
				if (size && Number.isFinite(Number(size.getAttribute('V'))))
					values['Char.Size'] = Number(size.getAttribute('V'));
			}
			shapes.push({
				id: shape.getAttribute('ID'),
				name: shape.getAttribute('NameU') ?? '',
				cells: values,
				colors,
			});
		}
		pages.push({ page: name, pageId: info.id, shapes });
	}
	return pages;
}
const { result: pages0, zip: zip0 } = await pageMap(corpus);
const pid = (name: string) => pages0.get(name)!.id;
const idOf = async (page: string, nameU: string) => {
	const doc = new DOMParser().parseFromString(
		await zip0.file(pages0.get(page)!.path)!.async('string'),
		'text/xml',
	);
	const hit = Array.from(doc.getElementsByTagName('Shape')).find(
		(s) => s.getAttribute('NameU') === nameU,
	);
	if (!hit) throw new Error(`No shape ${nameU} on ${page}`);
	return hit.getAttribute('ID')!;
};
const basicName = prefix + 'basic-shapes';
const basic = pid(basicName);
const rect = await idOf(basicName, 'Basic.Rect');
const oval = await idOf(basicName, 'Basic.Oval');
const line = await idOf(basicName, 'Basic.Line');
const group = await idOf(prefix + 'groups', 'Group.Main');
const groupsPage = pid(prefix + 'groups');
const containerPage = pid(flatDocument ? 'flat-glued-lines' : 'containers');
const connName = flatDocument ? 'flat-glued-lines' : 'connectors';
const connPage = pid(connName);
const connA = await idOf(
	flatDocument ? 'flat-glued-lines' : 'connectors',
	flatDocument ? 'Glue.Line' : 'Conn.A',
);

const glueName = flatDocument ? 'Glue.Line' : 'Dynamic connector';
const glueLine = await idOf(connName, glueName);
const glueTarget = await idOf(connName, flatDocument ? 'Glue.A' : 'Conn.A');
const glueOther = await idOf(connName, flatDocument ? 'Glue.B' : 'Conn.B');
// Master instances exist only in the full reference document.
const masterPage = flatDocument ? undefined : pid('master-instances');
const masterShape = async (name: string) => (masterPage ? idOf('master-instances', name) : '');

let captured: Awaited<ReturnType<typeof captureVisioClipboard>> | undefined;
let captureError = '';
try {
	captured = await captureVisioClipboard(corpus, basic, [rect, oval]);
} catch (error) {
	captureError = String(error);
}
// Multi-selection resize: the selection bounds (1..7 x 6..7.5) scale by 1.5 about the left-bottom
// corner, each member scaled and moved exactly as the viewer composes resize-shape and move-shape.
const k = 1.5;
const sel = { x0: 1, y0: 6 };
const scaleMember = (id: string, pinX: number, pinY: number, w: number, h: number): VisioEdit[] => [
	{ type: 'resize-shape', pageId: basic, shapeId: id, width: w * k, height: h * k },
	{
		type: 'move-shape',
		pageId: basic,
		shapeId: id,
		x: sel.x0 + (pinX - sel.x0) * k,
		y: sel.y0 + (pinY - sel.y0) * k,
	},
];
const cases: { name: string; family: string; edits: VisioEdit[] }[] = [
	{
		name: 'geometry-move-resize-rotate',
		family: 'geometry edits',
		edits: [
			{ type: 'move-shape', pageId: basic, shapeId: rect, x: 3, y: 5 },
			{ type: 'resize-shape', pageId: basic, shapeId: oval, width: 3, height: 2 },
			{ type: 'rotate-shape', pageId: basic, shapeId: rect, angle: Math.PI / 6 },
			{ type: 'flip-shape', pageId: basic, shapeId: oval, axis: 'horizontal' },
			{ type: 'move-line-endpoint', pageId: basic, shapeId: line, endpoint: 'end', x: 6, y: 3 },
		],
	},
	{
		name: 'geometry-anchored-resize',
		family: 'geometry edits',
		edits: [
			{
				type: 'resize-shape',
				pageId: basic,
				shapeId: rect,
				width: 4,
				height: 3,
				anchor: { x: 1, y: 0 },
			},
		],
	},
	{
		name: 'multi-selection-resize',
		family: 'group and multi-selection resize',
		edits: [...scaleMember(rect, 2.25, 6.75, 2.5, 1.5), ...scaleMember(oval, 5.75, 6.75, 2.5, 1.5)],
	},
	{
		name: 'group-resize',
		family: 'group and multi-selection resize',
		edits: [
			{
				type: 'resize-shape',
				pageId: groupsPage,
				shapeId: group,
				width: 5,
				height: 4,
				anchor: { x: 0, y: 0 },
			},
		],
	},
	{
		name: 'group-create',
		family: 'group and multi-selection resize',
		edits: [{ type: 'group-shapes', pageId: basic, shapeId: '50', memberIds: [rect, oval] }],
	},
	{
		name: 'group-ungroup',
		family: 'group and multi-selection resize',
		edits: [{ type: 'ungroup-shape', pageId: groupsPage, shapeId: group }],
	},
	{
		name: 'duplicate',
		family: 'duplication',
		edits: [
			{
				type: 'duplicate-shapes',
				pageId: basic,
				copies: [
					{ shapeId: rect, newShapeId: '60' },
					{ shapeId: oval, newShapeId: '61' },
				],
				offsetX: 0.5,
				offsetY: -1,
			},
		],
	},
	{
		name: 'duplicate-connected',
		family: 'duplication',
		edits: [
			{
				type: 'duplicate-shapes',
				pageId: connPage,
				copies: [{ shapeId: connA, newShapeId: '70' }],
				offsetX: 0,
				offsetY: flatDocument ? -2 : 2,
			},
		],
	},
	{
		name: 'clipboard-paste',
		family: 'clipboard paste',
		edits: [
			{
				type: 'paste-shapes',
				pageId: basic,
				clipboard: captured!,
				copies: [
					{ shapeId: rect, newShapeId: '80' },
					{ shapeId: oval, newShapeId: '81' },
				],
				offsetX: 1,
				offsetY: -3,
			},
		],
	},
	{
		name: 'formatting',
		family: 'formatting',
		edits: [
			{
				type: 'format-shape',
				pageId: basic,
				shapeId: rect,
				fillColor: '#336699',
				lineColor: '#cc0000',
				lineWeight: 0.04,
			},
			{
				type: 'format-text',
				pageId: basic,
				shapeId: rect,
				fontSize: 18,
				bold: true,
				horizontalAlign: 'left',
			},
			{
				type: 'format-shape',
				pageId: basic,
				shapeId: oval,
				linePattern: 2,
				fillPattern: 1,
				fillColor: '#ffcc00',
			},
		],
	},
	{
		name: 'arrangement',
		family: 'arrangement',
		edits: [{ type: 'reorder-shape', pageId: basic, shapeId: rect, order: 'front' }],
	},
	{
		name: 'container',
		family: 'containers',
		edits: [
			{
				type: 'insert-container',
				pageId: basic,
				shapeId: '90',
				memberIds: [rect, oval],
				style: 'classic',
				heading: 'Group A',
			},
		],
	},
	{
		name: 'container-move-with-members',
		family: 'containers',
		edits: [
			{
				type: 'insert-container',
				pageId: containerPage,
				shapeId: '91',
				memberIds: [],
				style: 'plain',
				heading: 'Empty',
				box: { x: 4, y: 2, width: 3, height: 2 },
			},
		],
	},
];

const masterProcess = await masterShape('Process');
const masterDecision = await masterShape('Decision');
const masterTerminator = await masterShape('Start/End');
// Edits of master instances (full document) and of glued lines and their shapes (both documents).
cases.push(
	...(masterPage
		? ([
				{
					name: 'master-instance-move',
					family: 'master instances',
					edits: [{ type: 'move-shape', pageId: masterPage, shapeId: masterDecision, x: 3, y: 8 }],
				},
				{
					name: 'master-instance-resize',
					family: 'master instances',
					edits: [
						{
							type: 'resize-shape',
							pageId: masterPage,
							shapeId: masterProcess,
							width: 2.5,
							height: 1.25,
						},
					],
				},
				{
					name: 'master-instance-format',
					family: 'master instances',
					edits: [
						{
							type: 'format-shape',
							pageId: masterPage,
							shapeId: masterTerminator,
							fillColor: '#336699',
							lineColor: '#cc0000',
							lineWeight: 2,
						},
						{
							type: 'format-text',
							pageId: masterPage,
							shapeId: masterTerminator,
							fontSize: 14,
							bold: true,
						},
					],
				},
				{
					name: 'master-instance-arrange',
					family: 'master instances',
					edits: [
						{ type: 'move-shape', pageId: masterPage, shapeId: masterProcess, x: 1.5, y: 7 },
						{
							type: 'resize-shape',
							pageId: masterPage,
							shapeId: masterDecision,
							width: 1.75,
							height: 1,
						},
						{ type: 'reorder-shape', pageId: masterPage, shapeId: masterProcess, order: 'front' },
					],
				},
			] as { name: string; family: string; edits: VisioEdit[] }[])
		: []),
	{
		name: 'glue-move-target',
		family: 'glue',
		edits: [{ type: 'move-shape', pageId: connPage, shapeId: glueTarget, x: 2, y: 3.75 }],
	},
	{
		name: 'glue-move-target-edge',
		family: 'glue',
		edits: [{ type: 'move-shape', pageId: connPage, shapeId: glueTarget, x: 2, y: 3 }],
	},
	{
		name: 'glue-move-both-targets',
		family: 'glue',
		edits: [
			{ type: 'move-shape', pageId: connPage, shapeId: glueTarget, x: 2.5, y: 7.5 },
			{ type: 'move-shape', pageId: connPage, shapeId: glueOther, x: 6.5, y: 2 },
		],
	},
	{
		name: 'glue-resize-target',
		family: 'glue',
		edits: [{ type: 'resize-shape', pageId: connPage, shapeId: glueTarget, width: 3, height: 0.8 }],
	},
	{
		name: 'glue-duplicate-line',
		family: 'glue',
		edits: [
			{
				type: 'duplicate-shapes',
				pageId: connPage,
				copies: [{ shapeId: glueLine, newShapeId: '75' }],
				offsetX: flatDocument ? 0.5 : -1,
				offsetY: flatDocument ? -1 : 3.5,
			},
		],
	},
	...[
		['overlap-a', 0.5, 2],
		['overlap-b', 3, 0],
		['overlap-c', 0.5, -1],
	].map(([suffix, offsetX, offsetY]) => ({
		name: 'glue-duplicate-line-' + suffix,
		family: 'glue',
		edits: [
			{
				type: 'duplicate-shapes' as const,
				pageId: connPage,
				copies: [{ shapeId: glueLine, newShapeId: '76' }],
				offsetX: offsetX as number,
				offsetY: offsetY as number,
			},
		],
	})),
	{
		name: 'glue-delete-line',
		family: 'glue',
		edits: [{ type: 'delete-shape', pageId: connPage, shapeId: glueLine }],
	},
);
const manifest: unknown[] = [];
for (const c of cases) {
	try {
		if (c.name === 'clipboard-paste' && !captured) throw new Error(captureError);
		const result = await editVsdx(corpus, c.edits);
		await writeFile(join(output, `${c.name}.vsdx`), result.bytes);
		manifest.push({
			name: c.name,
			family: c.family,
			status: 'written',
			pages: await claims(result.bytes),
		});
	} catch (error) {
		manifest.push({
			name: c.name,
			family: c.family,
			status: 'core-refused',
			error: String(error) + (process.env.STACK ? String((error as Error).stack) : ''),
		});
	}
}
await writeFile(join(output, 'cases.json'), JSON.stringify(manifest, null, 1));
console.log(output);
