// A minimal EditContext for testing command modules outside a full session.
import type { Workbook } from '../model';
import type { EditContext } from './context';
import { createCalcEngine } from './deps';
import { captureScope, type HistoryEntry, restoreSnapshot } from './history';
import type { WorkbookChangeKind } from './types';

export interface TestContext extends EditContext {
	readonly steps: { label: string; kind: WorkbookChangeKind; entries: HistoryEntry[] }[];
	/** Restores the last step's snapshots. */
	undo(): void;
}

export function testContext(workbook: Workbook): TestContext {
	const steps: TestContext['steps'] = [];
	return {
		workbook,
		calc: createCalcEngine(workbook),
		steps,
		run(label, kind, scopes, fn) {
			const befores = scopes.map((scope) => captureScope(workbook, scope));
			const result = fn(befores.find((snapshot) => snapshot.kind === 'refs')?.data.formulas);
			steps.push({
				label,
				kind,
				entries: befores.map((before, i) => ({
					before,
					after: captureScope(workbook, scopes[i] ?? { kind: 'meta' }, before),
				})),
			});
			return result;
		},
		undo() {
			const step = steps.pop();
			for (const entry of [...(step?.entries ?? [])].reverse())
				restoreSnapshot(workbook, entry.before);
		},
	};
}
