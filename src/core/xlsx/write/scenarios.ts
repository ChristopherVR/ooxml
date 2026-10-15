import type { ScenarioSet } from '../model';
import { attrs } from './xml-out';

/** `<scenarios>` from the model; empty when the sheet has no scenarios. */
export function scenariosXml(set: ScenarioSet | undefined): string {
	if (!set || !set.list.length) return '';
	const body = set.list
		.map((s) => {
			const cells = s.cells
				.map((c) => {
					const a = attrs({
						r: c.ref,
						deleted: c.deleted,
						undone: c.undone,
						val: c.value,
						numFmtId: c.numFmtId,
					});
					return `<inputCells${a}/>`;
				})
				.join('');
			const a = attrs({
				name: s.name,
				locked: s.locked || undefined,
				hidden: s.hidden || undefined,
				count: s.cells.length,
				user: s.user,
				comment: s.comment,
			});
			return `<scenario${a}>${cells}</scenario>`;
		})
		.join('');
	const a = attrs({ current: set.current, show: set.show, sqref: set.sqref });
	return `<scenarios${a}>${body}</scenarios>`;
}
