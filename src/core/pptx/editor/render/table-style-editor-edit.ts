/**
 * The write side of the table-style DEFINITION editor: a discriminated union
 * of field-level edits a binding's panel can make, and
 * {@link applyTableStyleFieldEdit}, the pure function that turns one such edit
 * into both an updated {@link ParsedTableStyleEntry} (for the binding to store
 * locally, e.g. React's `setTableStyleMap`) and a payload shaped exactly like
 * the `set_table_style_section` MCP tool's `SetTableStyleSectionParams`
 * (`packages/tools/src/tools/table-style-tools.ts`), so a host that DOES have
 * a `ToolContext` (the AI bridge) can run the identical mutation through that
 * tool instead of duplicating its merge logic.
 *
 * @module render/table-style-editor-edit
 */
import type {
	ParsedTableStyleBorder,
	ParsedTableStyleBorders,
	ParsedTableStyleEntry,
	ParsedTableStyleFill,
	ParsedTableStyleText,
	PptxTableCell3D,
	PptxThemeColorRef,
	PptxThemeFormatScheme,
} from 'ooxml-core/pptx';
import { resolveTableStyleLineRef, resolveTableStyleLineRefColor } from 'ooxml-core/pptx';

import { colorToFill } from './table-style-editor-descriptor';
import type { TableStyleBorderSide, TableStyleEditorPartId } from './table-style-editor-parts';
import { isTableStylePartName } from './table-style-editor-parts';

export type TableStyleEditorFieldEdit =
	| {
			readonly kind: 'fillColor';
			readonly hex: string;
			readonly ref: PptxThemeColorRef | undefined;
	  }
	| { readonly kind: 'fillNone'; readonly noFill: boolean }
	| { readonly kind: 'textBold'; readonly value: boolean }
	| { readonly kind: 'textItalic'; readonly value: boolean }
	| { readonly kind: 'textUnderline'; readonly value: boolean }
	| {
			readonly kind: 'textColor';
			readonly hex: string;
			readonly ref: PptxThemeColorRef | undefined;
	  }
	| {
			readonly kind: 'borderColor';
			readonly side: TableStyleBorderSide;
			readonly hex: string;
			readonly ref: PptxThemeColorRef | undefined;
	  }
	| { readonly kind: 'borderWidth'; readonly side: TableStyleBorderSide; readonly width: number }
	| { readonly kind: 'borderDash'; readonly side: TableStyleBorderSide; readonly dash: string }
	| { readonly kind: 'borderNone'; readonly side: TableStyleBorderSide; readonly noFill: boolean }
	| { readonly kind: 'cell3DBevelWidth'; readonly value: number | undefined }
	| { readonly kind: 'cell3DBevelHeight'; readonly value: number | undefined }
	| { readonly kind: 'cell3DBevelPreset'; readonly value: string | undefined };

/** Mirrors `packages/tools/src/tools/table-style-tools.ts`'s `SetTableStyleSectionParams`. */
export interface TableStyleSectionEditPayload {
	readonly styleId: string;
	readonly section: string;
	readonly fill?: ParsedTableStyleFill;
	readonly text?: ParsedTableStyleText;
	readonly borders?: ParsedTableStyleBorders;
	readonly cell3D?: PptxTableCell3D;
}

export interface TableStyleEditorEditResult {
	/** The full entry with the edit applied; store it back into `ParsedTableStyleMap[styleId]`. */
	readonly entry: ParsedTableStyleEntry;
	/** The SDK-shaped patch, for a caller that runs edits through `set_table_style_section` instead. */
	readonly payload: TableStyleSectionEditPayload;
}

/** Widen so `${part}Fill` / `${part}Text` / ... dynamic keys can be read AND assigned. */
type SectionKeyed = Record<
	string,
	| ParsedTableStyleFill
	| ParsedTableStyleText
	| ParsedTableStyleBorders
	| PptxTableCell3D
	| undefined
>;

/** What the edit path knows about the theme, for sides that reference it. */
export interface TableStyleEditContext {
	/**
	 * The theme's `a:fmtScheme`. With it, editing one facet of a theme line
	 * reference (`a:lnRef`) starts from the referenced line, so the facets the
	 * edit does not touch (width, dash) keep the theme's values.
	 */
	readonly formatScheme?: PptxThemeFormatScheme;
}

function patchBorderSide(
	borders: ParsedTableStyleBorders | undefined,
	side: TableStyleBorderSide,
	patch: Partial<ParsedTableStyleBorder>,
	formatScheme: PptxThemeFormatScheme | undefined,
): ParsedTableStyleBorders {
	const { lnRef, ...current } = borders?.[side] ?? {};
	if (!lnRef) {
		return { ...borders, [side]: { ...current, ...patch } };
	}
	const concrete = Object.keys(current).length > 0;
	const colourOnly =
		patch.fill !== undefined &&
		Object.keys(patch).every((key) => key === 'fill' || key === 'color');
	// A colour edit with no theme to resolve against keeps the reference and
	// swaps its colour, which is what the referenced line's `phClr` takes: the
	// theme line's width and dash survive the save. Turning it into a concrete
	// line here would drop them.
	if (colourOnly && !formatScheme && !concrete) {
		return { ...borders, [side]: { lnRef: { ...lnRef, color: patch.fill } } };
	}
	// Otherwise the side becomes a concrete line. Its starting point is the
	// referenced theme line when the scheme is known (width, dash, no-line, and
	// the colour it draws in: its own, or the reference's through its phClr),
	// else the reference's colour. An edit that sets a colour replaces it.
	const resolved = formatScheme
		? resolveTableStyleLineRef(lnRef, formatScheme, undefined)
		: undefined;
	const fromTheme: ParsedTableStyleBorder = {};
	if (resolved?.width !== undefined) {
		fromTheme.width = resolved.width;
	}
	if (resolved?.dash !== undefined) {
		fromTheme.dash = resolved.dash;
	}
	if (resolved?.noFill) {
		fromTheme.noFill = true;
	}
	const colour = resolveTableStyleLineRefColor(lnRef, formatScheme);
	const base: ParsedTableStyleBorder =
		colour && current.fill === undefined && current.color === undefined && !resolved?.noFill
			? { ...fromTheme, ...current, ...colour }
			: { ...fromTheme, ...current };
	return { ...borders, [side]: { ...base, ...patch } };
}

/**
 * Apply one field edit to `entry`'s `part`, returning the new entry plus the
 * SDK-shaped payload. `part === 'background'` only honours `fillColor` /
 * `fillNone` (background has no text/borders/cell3D facet); any other edit
 * kind on the background part is a no-op (returns `entry` unchanged, payload
 * with no facets set) so a caller does not need to pre-filter by part.
 */
export function applyTableStyleFieldEdit(
	entry: ParsedTableStyleEntry,
	part: TableStyleEditorPartId,
	edit: TableStyleEditorFieldEdit,
	context: TableStyleEditContext = {},
): TableStyleEditorEditResult {
	const { formatScheme } = context;
	if (part === 'background') {
		return applyBackgroundEdit(entry, edit);
	}
	const keyed = entry as unknown as SectionKeyed;
	const fillKey = `${part}Fill`;
	const textKey = `${part}Text`;
	const bordersKey = `${part}Borders`;
	const cell3DKey = `${part}Cell3D`;

	switch (edit.kind) {
		case 'fillColor':
			keyed[fillKey] = colorToFill(edit.hex, edit.ref);
			break;
		case 'fillNone': {
			const current = keyed[fillKey] as ParsedTableStyleFill | undefined;
			keyed[fillKey] = { schemeColor: current?.schemeColor ?? '', ...current, noFill: edit.noFill };
			break;
		}
		case 'textBold':
			keyed[textKey] = {
				...(keyed[textKey] as ParsedTableStyleText | undefined),
				bold: edit.value,
			};
			break;
		case 'textItalic':
			keyed[textKey] = {
				...(keyed[textKey] as ParsedTableStyleText | undefined),
				italic: edit.value,
			};
			break;
		case 'textUnderline':
			keyed[textKey] = {
				...(keyed[textKey] as ParsedTableStyleText | undefined),
				underline: edit.value,
			};
			break;
		case 'textColor': {
			const fill = colorToFill(edit.hex, edit.ref);
			keyed[textKey] = {
				...(keyed[textKey] as ParsedTableStyleText | undefined),
				fontSchemeColor: fill.schemeColor,
				fontTint: fill.tint,
				fontShade: fill.shade,
				fontColor: fill.color,
			};
			break;
		}
		case 'borderColor':
			keyed[bordersKey] = patchBorderSide(
				keyed[bordersKey] as ParsedTableStyleBorders | undefined,
				edit.side,
				{ fill: colorToFill(edit.hex, edit.ref), color: edit.ref ? undefined : edit.hex },
				formatScheme,
			);
			break;
		case 'borderWidth':
			keyed[bordersKey] = patchBorderSide(
				keyed[bordersKey] as ParsedTableStyleBorders | undefined,
				edit.side,
				{ width: edit.width },
				formatScheme,
			);
			break;
		case 'borderDash':
			keyed[bordersKey] = patchBorderSide(
				keyed[bordersKey] as ParsedTableStyleBorders | undefined,
				edit.side,
				{ dash: edit.dash },
				formatScheme,
			);
			break;
		case 'borderNone':
			keyed[bordersKey] = patchBorderSide(
				keyed[bordersKey] as ParsedTableStyleBorders | undefined,
				edit.side,
				{ noFill: edit.noFill },
				formatScheme,
			);
			break;
		case 'cell3DBevelWidth':
			keyed[cell3DKey] = {
				...(keyed[cell3DKey] as PptxTableCell3D | undefined),
				bevelWidth: edit.value,
			};
			break;
		case 'cell3DBevelHeight':
			keyed[cell3DKey] = {
				...(keyed[cell3DKey] as PptxTableCell3D | undefined),
				bevelHeight: edit.value,
			};
			break;
		case 'cell3DBevelPreset':
			keyed[cell3DKey] = {
				...(keyed[cell3DKey] as PptxTableCell3D | undefined),
				bevelPreset: edit.value,
			};
			break;
	}

	const nextEntry = keyed as unknown as ParsedTableStyleEntry;
	return {
		entry: nextEntry,
		payload: {
			styleId: entry.styleId,
			section: part,
			...(isTableStylePartName(part)
				? {
						fill: keyed[fillKey] as ParsedTableStyleFill | undefined,
						text: keyed[textKey] as ParsedTableStyleText | undefined,
						borders: keyed[bordersKey] as ParsedTableStyleBorders | undefined,
						cell3D: keyed[cell3DKey] as PptxTableCell3D | undefined,
					}
				: {}),
		},
	};
}

function applyBackgroundEdit(
	entry: ParsedTableStyleEntry,
	edit: TableStyleEditorFieldEdit,
): TableStyleEditorEditResult {
	const current = entry.tableBackground;
	let nextBackground = current;
	if (edit.kind === 'fillColor') {
		nextBackground = { ...current, fill: colorToFill(edit.hex, edit.ref), fillRef: undefined };
	} else if (edit.kind === 'fillNone') {
		nextBackground = {
			...current,
			fill: {
				schemeColor: current?.fill?.schemeColor ?? '',
				...current?.fill,
				noFill: edit.noFill,
			},
		};
	}
	const nextEntry: ParsedTableStyleEntry = { ...entry, tableBackground: nextBackground };
	return {
		entry: nextEntry,
		payload: { styleId: entry.styleId, section: 'background' },
	};
}
