// SmartArt editing in a worksheet: the shared node commands (`applyDiagramNodeEdit` of
// `ooxml-core/diagram`, which Word uses too), Change Layout and Insert SmartArt, each one undoable
// step of the edit session. An edit drops the cached drawing: Excel never paints it and lays an
// edited diagram out again from its data and layout parts (evidence in
// docs/agnostic-core-plan.md, "Excel and an edited data part"), so the save removes it too.
import {
	applyDiagramNodeEdit,
	createDiagramNodes,
	formatDiagramSource,
	type DiagramFormatChange,
	type DiagramEditOptions,
	type DiagramLayoutType,
	type DiagramNodeEdit,
} from '../../diagram/index';
import type { DrawingAnchor, SmartArtObject } from '../model';
import { SMART_ART_NOTICES } from '../smart-art-notices';
import { type EditContext, sheetAt, sheetDrawingAt } from './context';

export interface SmartArtEditResult {
	changed: boolean;
	/** The node the command created, or the node it acted on (for the text pane's focus). */
	nodeId?: string;
}

/** A new layout for Change Layout: its definition XML comes from `ooxml-core/diagram/builtin`. */
export interface SmartArtLayoutChange {
	layoutId: string;
	family?: DiagramLayoutType | undefined;
	layoutXml: string;
}

/** What Insert SmartArt needs: the definitions come from `ooxml-core/diagram/builtin`. */
export interface SmartArtInsertion extends SmartArtLayoutChange {
	anchor: DrawingAnchor;
	colorsId: string;
	colorsXml: string;
	quickStyleId: string;
	quickStyleXml: string;
	/** The node texts the diagram starts with (a hierarchy puts the others under the first). */
	items: readonly string[];
	name?: string;
}

function smartArtAt(ctx: EditContext, s: number, index: number): SmartArtObject {
	const drawing = sheetDrawingAt(ctx.workbook, s, index);
	if (drawing?.kind !== 'smartArt')
		throw new RangeError(`No SmartArt graphic at index ${index} on sheet ${s}`);
	return drawing;
}

/** Marks a diagram edited: its cached drawing is stale, so it is dropped (and on save). */
function markEdited(object: SmartArtObject): void {
	delete object.diagram;
	delete object.cacheGenerated;
	if (!object.inserted) object.edited = true;
	object.notice = SMART_ART_NOTICES.edited;
	object.issues = object.issues.filter((issue) => !issue.code.startsWith('DIAGRAM_DRAWING'));
}

export function formatSmartArt(
	ctx: EditContext,
	s: number,
	index: number,
	change: DiagramFormatChange,
): boolean {
	const object = smartArtAt(ctx, s, index);
	if (
		(change.kind === 'colors' ? object.colorsId : object.quickStyleId) ===
		`urn:ooxml-core/${change.kind}/${change.value}`
	)
		return false;
	const next = formatDiagramSource(
		{
			...object.layoutSource,
			...(object.definitions?.colors ? { colorsXml: object.definitions.colors } : {}),
			...(object.definitions?.quickStyle ? { quickStyleXml: object.definitions.quickStyle } : {}),
		},
		change,
	);
	ctx.run(
		'Format SmartArt',
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['drawings'] }],
		() => {
			const live = smartArtAt(ctx, s, index);
			live.layoutSource = next.source;
			live.definitions = {
				...live.definitions,
				...(change.kind === 'colors'
					? { colors: next.source.colorsXml! }
					: { quickStyle: next.source.quickStyleXml! }),
			};
			if (next.colorsId) live.colorsId = next.colorsId;
			if (next.quickStyleId) live.quickStyleId = next.quickStyleId;
			markEdited(live);
		},
		{ sheet: s },
	);
	return true;
}

/** Applies one SmartArt Design or text pane command (one undo step when it changes anything). */
export function editSmartArt(
	ctx: EditContext,
	s: number,
	index: number,
	edit: DiagramNodeEdit,
	options?: DiagramEditOptions,
): SmartArtEditResult {
	const object = smartArtAt(ctx, s, index);
	const result = applyDiagramNodeEdit(object.nodes, edit, options);
	const nodeId = result.nodeId ? { nodeId: result.nodeId } : {};
	if (!result.changed) return { changed: false, ...nodeId };
	ctx.run(
		edit.kind === 'text' ? 'Typing' : 'Edit SmartArt',
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['drawings'] }],
		() => {
			const live = smartArtAt(ctx, s, index);
			live.nodes = result.nodes;
			markEdited(live);
		},
		{ sheet: s },
	);
	return { changed: true, ...nodeId };
}

/** SmartArt Design > Change Layout; false when the diagram already uses the layout. */
export function changeSmartArtLayout(
	ctx: EditContext,
	s: number,
	index: number,
	change: SmartArtLayoutChange,
): boolean {
	const object = smartArtAt(ctx, s, index);
	if (object.layoutId === change.layoutId) return false;
	ctx.run(
		'Change SmartArt layout',
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['drawings'] }],
		() => {
			const live = smartArtAt(ctx, s, index);
			live.layoutId = change.layoutId;
			if (change.family) live.layoutFamily = change.family;
			else delete live.layoutFamily;
			live.definitions = { ...live.definitions, layout: change.layoutXml };
			markEdited(live);
		},
		{ sheet: s },
	);
	return true;
}

/** Insert > SmartArt: adds a new diagram (its parts are written on save); returns its index. */
export function insertSmartArt(
	ctx: EditContext,
	s: number,
	insertion: SmartArtInsertion,
	options?: DiagramEditOptions,
): number {
	const sheet = sheetAt(ctx.workbook, s);
	const object: SmartArtObject = {
		kind: 'smartArt',
		anchor: structuredClone(insertion.anchor),
		name: insertion.name ?? 'Diagram',
		nodes: createDiagramNodes(insertion.items, insertion.family, options),
		layoutId: insertion.layoutId,
		...(insertion.family ? { layoutFamily: insertion.family } : {}),
		colorsId: insertion.colorsId,
		quickStyleId: insertion.quickStyleId,
		definitions: {
			layout: insertion.layoutXml,
			colors: insertion.colorsXml,
			quickStyle: insertion.quickStyleXml,
		},
		inserted: true,
		notice: SMART_ART_NOTICES.edited,
		issues: [],
		sourceXml: '',
	};
	return ctx.run(
		'Insert SmartArt',
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['drawings'] }],
		() => {
			sheet.drawings.push(object);
			return sheet.drawings.length - 1;
		},
		{ sheet: s },
	);
}
