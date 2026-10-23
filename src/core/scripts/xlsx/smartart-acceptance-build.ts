// Builds the workbooks `smartart-excel-acceptance.ps1` opens in Excel. Run with bun from src/core:
//   bun scripts/xlsx/smartart-acceptance-build.ts <output folder>
// Every file is written by `saveXlsx`: the committed Excel SmartArt fixture saved untouched, after
// node edits (text, add, remove, move, promote, demote), after Change Layout, and new workbooks
// with SmartArt inserted from Insert SmartArt presets. `manifest.json` lists, per file, the node
// texts Excel must report for each SmartArt (`SmartArt.AllNodes`, tree order) and its layout id.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
	BUILTIN_SMARTART_COLORS_ID,
	BUILTIN_SMARTART_COLORS_XML,
	BUILTIN_SMARTART_QUICK_STYLE_ID,
	BUILTIN_SMARTART_QUICK_STYLE_XML,
	builtinLayoutIdForPreset,
	loadBuiltinSmartArtLayoutXml,
} from '../../diagram/builtin/index';
import { DIAGRAM_INSERT_PRESETS } from '../../diagram/catalogue';
import { createEditSession, type EditSession } from '../../xlsx/edit/index';
import type { SmartArtObject, Workbook } from '../../xlsx/model';
import { loadXlsx } from '../../xlsx/read/index';
import { createWorkbook } from '../../xlsx/workbook';
import { saveXlsx } from '../../xlsx/write/index';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/xlsx/smartart-acceptance-build.ts <output folder>');
mkdirSync(out, { recursive: true });
const fixture = join(import.meta.dir, '..', '..', 'xlsx', '__fixtures__', 'excel-smartart.xlsx');
const manifest: Record<string, { texts: string[]; layout?: string }[]> = {};

const smartArts = (wb: Workbook) =>
	wb.sheets.flatMap((sheet) =>
		sheet.drawings.filter((d): d is SmartArtObject => d.kind === 'smartArt'),
	);

/** Excel's AllNodes order is the tree order the model keeps; empty nodes report ''. */
const expected = (wb: Workbook) =>
	smartArts(wb).map((art) => ({
		texts: art.nodes.map((node) => node.text),
		...(art.layoutId ? { layout: art.layoutId } : {}),
	}));

async function write(name: string, wb: Workbook): Promise<void> {
	writeFileSync(join(out!, name), await saveXlsx(wb));
	manifest[name] = expected(wb);
	console.log(name);
}

async function edited(name: string, edit: (session: EditSession, art: SmartArtObject) => void) {
	const wb = await loadXlsx(new Uint8Array(readFileSync(fixture)));
	const session = createEditSession(wb);
	edit(session, smartArts(wb)[0]!);
	await write(name, wb);
}

const index = (session: EditSession, art: SmartArtObject) =>
	(session.workbook.sheets[0]?.drawings ?? []).indexOf(art);
const ids = (art: SmartArtObject) => art.nodes.map((node) => node.id);

const layout = (preset: string) => {
	const layoutId = builtinLayoutIdForPreset(preset);
	if (!layoutId) throw new Error(`no built-in layout for ${preset}`);
	return { layoutId, layoutXml: loadBuiltinSmartArtLayoutXml(layoutId) };
};

await write('smartart-untouched.xlsx', await loadXlsx(new Uint8Array(readFileSync(fixture))));

await edited('smartart-edited.xlsx', (session, art) => {
	const [plan, build, ship] = ids(art);
	const at = index(session, art);
	session.editSmartArt(0, at, { kind: 'text', nodeId: build!, text: 'Build & test' });
	session.editSmartArt(0, at, { kind: 'add', position: 'after', nodeId: ship, text: 'Support' });
	session.editSmartArt(0, at, { kind: 'move', nodeId: plan!, direction: 1 });
});

await edited('smartart-removed.xlsx', (session, art) => {
	session.editSmartArt(0, index(session, art), { kind: 'remove', nodeId: ids(art)[1]! });
});

await edited('smartart-promoted-demoted.xlsx', (session, art) => {
	const at = index(session, art);
	const [, build, ship] = ids(art);
	session.editSmartArt(0, at, { kind: 'add', position: 'below', nodeId: build, text: 'Child' });
	session.editSmartArt(0, at, { kind: 'demote', nodeId: ship! });
	session.editSmartArt(0, at, { kind: 'promote', nodeId: ship! });
	session.editSmartArt(0, at, { kind: 'demote', nodeId: ship! });
});

await edited('smartart-layout-changed.xlsx', (session, art) => {
	session.changeSmartArtLayout(0, index(session, art), {
		...layout('basicChevronProcess'),
		family: 'process',
	});
});

{
	const wb = createWorkbook();
	const session = createEditSession(wb);
	const presets = ['basicBlockList', 'basicChevronProcess', 'basicCycle', 'hierarchy'];
	presets.forEach((name, n) => {
		const preset = DIAGRAM_INSERT_PRESETS.find((entry) => entry.layout === name);
		if (!preset) throw new Error(`no preset ${name}`);
		session.insertSmartArt(0, {
			anchor: {
				from: { row: n * 16, col: 1, rowOffset: 0, colOffset: 0 },
				ext: { cx: 4572000, cy: 2667000 },
			},
			...layout(name),
			family: preset.family,
			colorsId: BUILTIN_SMARTART_COLORS_ID,
			colorsXml: BUILTIN_SMARTART_COLORS_XML,
			quickStyleId: BUILTIN_SMARTART_QUICK_STYLE_ID,
			quickStyleXml: BUILTIN_SMARTART_QUICK_STYLE_XML,
			items: preset.defaultItems,
		});
	});
	await write('smartart-inserted.xlsx', wb);
}

writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, '\t'));
