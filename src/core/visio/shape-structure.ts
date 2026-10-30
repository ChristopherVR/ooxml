import type { VisioContainerOptions, VisioShapeStructure, VisioSwimlaneRole } from './model';
import { sectionRows, type Sheet } from './sheet';
import {
	VISIO_CALLOUT_LEADER_ROW,
	VISIO_CALLOUT_TARGET_ROW,
	VISIO_CONTAINER_MEMBERS_ROW,
} from './edit-diagram-parts-commands';
import {
	SHEET_ID as ID,
	VISIO_RELATIONSHIP_CALLOUT_TARGET,
	VISIO_RELATIONSHIP_MEMBERS,
	dependsOn,
	sheetIdList,
} from './container-relationships';
export type { VisioShapeStructure } from './model';

/** Strip the quotes Visio writes around string results of formulas. */
const unquote = (value: string | undefined) => value?.replace(/^"(.*)"$/s, '$1').trim();

/**
 * Container Format options stored as local User rows (margin, resize mode, lock). Absent rows
 * mean Visio's defaults, so only the rows that exist are reported.
 */
function containerOptions(
	rows: ReadonlyMap<string, string | undefined>,
): VisioContainerOptions | undefined {
	const options: VisioContainerOptions = {};
	const margin = Number(unquote(rows.get('msvSDContainerMargin')));
	if (unquote(rows.get('msvSDContainerMargin')) !== undefined && Number.isFinite(margin))
		options.margin = margin;
	const resize = Number(unquote(rows.get('msvSDContainerResize')));
	if (resize === 0 || resize === 1 || resize === 2) options.resize = resize;
	const locked = unquote(rows.get('msvSDContainerLocked'))?.toUpperCase();
	if (locked !== undefined) options.locked = locked === '1' || locked === 'TRUE';
	return Object.keys(options).length ? options : undefined;
}

/** Container members from the editor's older private User row, else from Relationships text. */
export function containerMemberIds(
	ownRow: string | undefined,
	relationships: string | undefined,
): string[] {
	return [
		...new Set(
			ownRow !== undefined
				? sheetIdList(unquote(ownRow))
				: dependsOn(relationships, VISIO_RELATIONSHIP_MEMBERS),
		),
	];
}

/** The editor's swimlane marker rows (non-native: Visio keeps its lists in master formulas). */
export const VISIO_SWIMLANE_FRAME_ROW = 'ooxmlSwimlaneFrame';
export const VISIO_SWIMLANE_LANE_ROW = 'ooxmlSwimlane';
function swimlaneRole(
	rows: ReadonlyMap<string, string | undefined>,
): VisioSwimlaneRole | undefined {
	if (unquote(rows.get(VISIO_SWIMLANE_LANE_ROW)) === '1') return { role: 'lane' };
	if (unquote(rows.get(VISIO_SWIMLANE_FRAME_ROW)) !== '1') return undefined;
	const flag = (name: string, fallback: boolean) => {
		const value = unquote(rows.get(name));
		return value === undefined ? fallback : value !== '0';
	};
	const style = Number(unquote(rows.get('visCFFStyle')));
	return {
		role: 'frame',
		options: {
			direction: flag('CFFVertical', false) ? 'vertical' : 'horizontal',
			showTitle: flag('visShowTitle', true),
			showSeparators: flag('visShowPhase', true),
			rotateLabels: flag('visRotateLabel', false),
			style: Number.isInteger(style) && style >= 1 && style <= 4 ? style : 1,
		},
	};
}

/**
 * Read a resolved sheet's User.msvStructureType. Members come from Visio's Relationships cell
 * (DEPENDSON 4 on a container, 6 for a callout's target), read as text only: the formulas are
 * never run. The editor's older private User row is still read first, for files it saved before
 * it wrote Relationships. Lists and other structures are not modelled.
 */
export function visioShapeStructure(sheet: Sheet): VisioShapeStructure | undefined {
	const rows = new Map(
		sectionRows(sheet, 'User').map((row) => [row.name ?? '', row.cells.get('Value')?.value]),
	);
	const type = unquote(rows.get('msvStructureType'))?.toLowerCase();
	const relationships = sheet.cells.get('Relationships')?.formula;
	if (type === 'container') {
		const swimlane = swimlaneRole(rows);
		const options = containerOptions(rows);
		return {
			type,
			memberIds: containerMemberIds(rows.get(VISIO_CONTAINER_MEMBERS_ROW), relationships),
			...(swimlane ? { swimlane } : {}),
			...(options ? { options } : {}),
		};
	}
	if (type === 'callout') {
		const target =
			unquote(rows.get(VISIO_CALLOUT_TARGET_ROW)) ??
			dependsOn(relationships, VISIO_RELATIONSHIP_CALLOUT_TARGET)[0];
		const leader = unquote(rows.get(VISIO_CALLOUT_LEADER_ROW));
		return {
			type,
			...(target && ID.test(target) ? { targetId: target } : {}),
			...(leader && ID.test(leader) ? { leaderId: leader } : {}),
		};
	}
	return undefined;
}
