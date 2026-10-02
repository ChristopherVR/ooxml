// A minimal EditContext for testing command modules outside a full session.
import type { Workbook } from '../model.js';
import type { EditContext } from './context.js';
import { createCalcEngine } from './deps.js';
import { captureScope, type HistoryEntry, restoreSnapshot } from './history.js';
import type { WorkbookChangeKind } from './types.js';

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
			const result = fn();
			steps.push({
				label,
				kind,
				entries: befores.map((before, i) => ({
					before,
					after: captureScope(workbook, scopes[i] ?? { kind: 'meta' }),
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
