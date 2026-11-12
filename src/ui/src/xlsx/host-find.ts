// Find and Replace for a host's own UI. Finding never needs edit rights; replacing re-reads
// read-only state and sheet protection and only touches cells the user could type into.
import { effectiveStyleId, styleAt, type FindMatch } from 'ooxml-core/xlsx';
import { stepMatch } from 'ooxml-core/xlsx/ui';
import type { EditorCore } from './editor-core';
import { selectOn } from './dialogs/go-to';
import { canEditNow, commitPending } from './host-workflows';
import type { FindOperations } from './surface-workflows';

export function findOperations(core: EditorCore): FindOperations {
	const ctx = core.ctx;
	const available = (): boolean => Boolean(core.host.isConnected && core.session && !core.loading);
	/** A cell on a protected sheet is editable only when it is not locked. */
	const replaceable = (match: FindMatch): boolean => {
		const workbook = core.workbook;
		const ws = workbook?.sheets[match.sheet];
		if (!workbook || !ws) return false;
		if (!ws.protection?.sheet) return true;
		const style = styleAt(workbook, effectiveStyleId(ws, match.row, match.col));
		return style.protection?.locked === false;
	};
	const at = (match: FindMatch): boolean =>
		match.sheet === core.activeSheet &&
		match.row === core.selection.get().active.row &&
		match.col === core.selection.get().active.col;
	const select = (match: FindMatch): boolean => {
		if (!available() || !core.workbook?.sheets[match.sheet]) return false;
		selectOn(ctx, match.sheet, { start: match, end: match });
		core.requestRender();
		return true;
	};
	const findAll: FindOperations['findAll'] = (query) =>
		available() && query.text ? (core.session?.findAll(query) ?? []) : [];
	const findNext: FindOperations['findNext'] = (query, options) => {
		if (!commitPending(core)) return undefined;
		const active = core.selection.get().active;
		const match = stepMatch(
			findAll(query),
			core.activeSheet,
			active.row,
			active.col,
			query.order === 'columns',
			options?.backwards === true,
		);
		return match && select(match) ? match : undefined;
	};
	return {
		canReplace: () => canEditNow(core),
		findAll,
		findNext,
		select,
		replace(query, replacement) {
			if (!canEditNow(core) || !commitPending(core) || !canEditNow(core)) return false;
			const current = findAll(query).find((match) => at(match) && replaceable(match));
			const replaced = current
				? Boolean(core.session?.replaceOne(query, replacement, current))
				: false;
			if (replaced) core.requestRender();
			findNext(query);
			return replaced;
		},
		replaceAll(query, replacement) {
			if (!canEditNow(core) || !commitPending(core) || !canEditNow(core)) return 0;
			const session = core.session;
			if (!session) return 0;
			const matches = findAll(query);
			let count = 0;
			if (matches.every(replaceable)) count = session.replaceAll(query, replacement);
			else
				session.batch('Replace All', () => {
					for (const match of matches)
						if (replaceable(match) && session.replaceOne(query, replacement, match)) count++;
				});
			if (count) core.requestRender();
			return count;
		},
	};
}
