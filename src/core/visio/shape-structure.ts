import type { VisioShapeStructure, VisioSwimlaneRole } from './model';
import { sectionRows, type Sheet } from './sheet';
import {
	VISIO_CALLOUT_LEADER_ROW,
	VISIO_CALLOUT_TARGET_ROW,
	VISIO_CONTAINER_MEMBERS_ROW,
} from './edit-diagram-parts-commands';
import {
	SHEET_ID as ID,
	VISIO_RELATIONSHIP_CALLOUT_TARGET,
	VISIO_RELATIONSHIP_LIST_MEMBERS,
	VISIO_RELATIONSHIP_MEMBERS,
	VISIO_RELATIONSHIP_MEMBER_OF,
	dependsOn,
	relationTerms,
	sheetIdList,
} from './container-relationships';
export type { VisioShapeStructure } from './model';

/** Strip the quotes Visio writes around string results of formulas. */
const unquote = (value: string | undefined) => value?.replace(/^"(.*)"$/s, '$1').trim();

/**
 * Container members from the editor's older private User row, else from Relationships text
 * (kind 1). A formula with a kind 4 term and no kind 1 term is the form this editor wrote before
 * 2026-10-30, which put the members under 4.
 */
export function containerMemberIds(
	ownRow: string | undefined,
	relationships: string | undefined,
): string[] {
	if (ownRow !== undefined) return [...new Set(sheetIdList(unquote(ownRow)))];
	const terms = relationTerms(relationships);
	const legacy = terms && !terms.some((term) => term.kind === VISIO_RELATIONSHIP_MEMBERS);
	return [
		...new Set(
			dependsOn(relationships, legacy ? VISIO_RELATIONSHIP_MEMBER_OF : VISIO_RELATIONSHIP_MEMBERS),
		),
	];
}

/**
 * The swimlane markers. Lanes, frames and separators carry Visio's own User.msvShapeCategories
 * (Swimlane, CFF Container, Phase) and are members of its Swimlane List and Phase List. The
 * editor's own marker rows are written too and still read, for flowcharts it saved as plain
 * containers.
 */
export const VISIO_SWIMLANE_FRAME_ROW = 'ooxmlSwimlaneFrame';
export const VISIO_SWIMLANE_LANE_ROW = 'ooxmlSwimlane';
export const VISIO_SEPARATOR_ROW = 'ooxmlSeparator';
const categoriesOf = (rows: ReadonlyMap<string, string | undefined>): string[] =>
	(unquote(rows.get('msvShapeCategories')) ?? '').split(';').map((item) => item.trim());
function swimlaneRole(
	rows: ReadonlyMap<string, string | undefined>,
): VisioSwimlaneRole | undefined {
	const categories = categoriesOf(rows);
	if (unquote(rows.get(VISIO_SWIMLANE_LANE_ROW)) === '1' || categories.includes('Swimlane'))
		return { role: 'lane' };
	if (unquote(rows.get(VISIO_SEPARATOR_ROW)) === '1' || categories.includes('Phase'))
		return { role: 'separator' };
	if (unquote(rows.get(VISIO_SWIMLANE_FRAME_ROW)) !== '1' && !categories.includes('CFF Container'))
		return undefined;
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
		return {
			type,
			memberIds: containerMemberIds(rows.get(VISIO_CONTAINER_MEMBERS_ROW), relationships),
			...(swimlane ? { swimlane } : {}),
		};
	}
	if (type === 'list') {
		const required = unquote(rows.get('msvSDListRequiredCategories'));
		const direction = Number(unquote(rows.get('msvSDListDirection')));
		return {
			type,
			memberIds: [...new Set(dependsOn(relationships, VISIO_RELATIONSHIP_LIST_MEMBERS))],
			direction: Number.isInteger(direction) ? direction : 0,
			...(required === 'Swimlane' ? { kind: 'swimlane' as const } : {}),
			...(required === 'Phase' ? { kind: 'phase' as const } : {}),
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
