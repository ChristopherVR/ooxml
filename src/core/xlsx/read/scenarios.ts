// Scenario Manager scenarios (`<scenarios>` in a worksheet).
import type { XmlElement } from '../../xml/index';
import type { Scenario, ScenarioCell, ScenarioSet } from '../model';
import { att, boolAttr, numAttr, xChildren } from './xml-util';

export function readScenarios(node: XmlElement | undefined): ScenarioSet | undefined {
	if (!node) return undefined;
	const list: Scenario[] = [];
	for (const el of xChildren(node, 'scenario')) {
		const scenario: Scenario = { name: att(el, 'name') ?? '', cells: [] };
		if (boolAttr(el, 'locked')) scenario.locked = true;
		if (boolAttr(el, 'hidden')) scenario.hidden = true;
		const user = att(el, 'user');
		if (user) scenario.user = user;
		const comment = att(el, 'comment');
		if (comment) scenario.comment = comment;
		for (const input of xChildren(el, 'inputCells')) {
			const cell: ScenarioCell = { ref: att(input, 'r') ?? '', value: att(input, 'val') ?? '' };
			if (boolAttr(input, 'deleted')) cell.deleted = true;
			if (boolAttr(input, 'undone')) cell.undone = true;
			const numFmtId = numAttr(input, 'numFmtId');
			if (numFmtId !== undefined) cell.numFmtId = numFmtId;
			if (cell.ref) scenario.cells.push(cell);
		}
		list.push(scenario);
	}
	const set: ScenarioSet = { list };
	const current = numAttr(node, 'current');
	if (current !== undefined) set.current = current;
	const show = numAttr(node, 'show');
	if (show !== undefined) set.show = show;
	const sqref = att(node, 'sqref');
	if (sqref) set.sqref = sqref;
	return set;
}
