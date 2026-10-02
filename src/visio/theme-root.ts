/**
 * MS-VSDX 2.4.4.58: ColorSchemeIndex 0 selects the root style's colors.
 * MS-VSDX 2.4.4.70 likewise specifies root formats when ColorSchemeIndex is zero.
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/9753d977-a1de-49e8-888c-cf532efe7982
 * MS-VSDX 2.4.4.275: QuickStyleLineMatrix 0 selects root line properties.
 * MS-VSDX 2.4.4.271: QuickStyleFillMatrix 0 selects root fill properties.
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/25689058-b1e7-4d3c-a833-0a4c7180f5f2
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/1e7e9b7e-d116-41c0-9c53-5e57c26042a4
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/edfd9f33-fea3-4bf5-8cd5-a91a3d677b03
 * Resolve saved selectors only. Missing or malformed selectors are not root selection.
 */
import type { Cells, Report, Sheet } from './sheet.js';
import { integer } from './theme-color.js';
import { rootThemeSelected, type ThemeResources } from './theme-resolve.js';

const LINE_PROPERTIES = [
	'LineWeight',
	'LineCap',
	'LinePattern',
	'LineColorTrans',
	'BeginArrow',
	'EndArrow',
	'BeginArrowSize',
	'EndArrowSize',
];
// Scalar fill formats only. Saved gradient stop rows still require a separate renderer.
const FILL_PROPERTIES = [
	'FillForegndTrans',
	'FillBkgndTrans',
	'FillPattern',
	'FillGradientDir',
	'FillGradientAngle',
	'FillGradientEnabled',
	'RotateGradientWithShape',
	'UseGroupGradient',
];
// Parsed root sheets are immutable during normalization. Cache missing rows as well.
const rootCharacters = new WeakMap<Sheet, Cells | undefined>();
function rootCharacterCells(root: Sheet): Cells | undefined {
	if (rootCharacters.has(root)) return rootCharacters.get(root);
	for (const section of root.sections.values()) {
		if (section.name !== 'Character' || section.deleted) continue;
		for (const row of section.rows.values()) {
			if (row.index !== '0' || row.deleted) continue;
			rootCharacters.set(root, row.cells);
			return row.cells;
		}
	}
	rootCharacters.set(root, undefined);
	return undefined;
}
function replaceThemed(
	cells: Cells,
	root: Cells,
	names: readonly string[],
	report?: Report,
): Cells {
	let result = cells;
	for (const name of names) {
		const saved = root.get(name);
		if (cells.get(name)?.value !== 'Themed' || !saved?.value || saved.value === 'Themed') continue;
		if (result === cells) result = new Map(cells);
		result.set(name, saved);
		// Original-cell errors were already reported before substitution. Do not duplicate them.
		if (saved.error !== undefined && cells.get(name)?.error === undefined)
			report?.(
				'cached-cell-error',
				`Cell ${name} records a formula error; only its last valid cached value is available.`,
			);
	}
	return result;
}
/** Never mutate inherited sheets or fill absent caches from a root-style guess. */
export function rootThemeSheet(sheet: Sheet, resources: ThemeResources, report?: Report): Sheet {
	const root = resources.rootSheet;
	if (!root) return sheet;
	const colorRoot = rootThemeSelected(sheet.cells, resources);
	const lineRoot = colorRoot || integer(sheet.cells.get('QuickStyleLineMatrix')?.value) === 0;
	const fillRoot = colorRoot || integer(sheet.cells.get('QuickStyleFillMatrix')?.value) === 0;
	let cells = sheet.cells;
	if (colorRoot)
		cells = replaceThemed(cells, root.cells, ['LineColor', 'FillForegnd', 'FillBkgnd'], report);
	if (lineRoot) cells = replaceThemed(cells, root.cells, LINE_PROPERTIES, report);
	if (fillRoot) cells = replaceThemed(cells, root.cells, FILL_PROPERTIES, report);
	let sections = sheet.sections;
	const rootCharacter = colorRoot ? rootCharacterCells(root) : undefined;
	if (colorRoot && rootCharacter) {
		for (const [key, section] of sheet.sections) {
			if (section.name !== 'Character' || section.deleted) continue;
			const rows = new Map(section.rows);
			for (const [index, row] of rows) {
				if (row.deleted) continue;
				const resolved = replaceThemed(row.cells, rootCharacter, ['Color'], report);
				if (resolved !== row.cells) rows.set(index, { ...row, cells: resolved });
			}
			if (sections === sheet.sections) sections = new Map(sheet.sections);
			sections.set(key, { ...section, rows });
		}
	}
	return cells === sheet.cells && sections === sheet.sections ? sheet : { cells, sections };
}
