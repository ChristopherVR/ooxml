import type { VisioDiagnostic } from './model.js';
import type { Report } from './sheet.js';

const MAX_CHARACTERS = 1_000_000;
const marker: VisioDiagnostic = {
	code: 'diagnostics-truncated',
	severity: 'warning',
	message: 'Additional diagnostic content was omitted by the configured count or character limits.',
};
/** Bound retained strings before constructing deduplication keys. */
export function diagnosticCollector(maxCount: number): {
	report: Report;
	finish: () => VisioDiagnostic[];
} {
	const diagnostics: VisioDiagnostic[] = [],
		keys = new Set<string>();
	let characters = 0,
		truncated = false;
	const bounded = (value: string, max: number) => {
		if (value.length <= max) return value;
		truncated = true;
		return value.slice(0, max - 1) + '…';
	};
	const report: Report = (code, message, context = {}) => {
		if (diagnostics.length >= maxCount) {
			truncated = true;
			return;
		}
		const item: VisioDiagnostic = {
			code: bounded(code, 64),
			severity: context.severity ?? 'warning',
			message: bounded(message, 2048),
			...(context.part === undefined ? {} : { part: bounded(context.part, 1024) }),
			...(context.pageId === undefined ? {} : { pageId: bounded(context.pageId, 256) }),
			...(context.shapeId === undefined ? {} : { shapeId: bounded(context.shapeId, 1024) }),
		};
		const key = JSON.stringify(item);
		if (keys.has(key)) return;
		const size = Object.values(item).reduce((sum, value) => sum + value.length, 0);
		if (characters + size > MAX_CHARACTERS - 256) {
			truncated = true;
			return;
		}
		characters += size;
		keys.add(key);
		diagnostics.push(item);
	};
	return {
		report,
		finish: () => {
			if (truncated) {
				if (diagnostics.length >= maxCount) diagnostics.pop();
				diagnostics.push({ ...marker });
			}
			return diagnostics;
		},
	};
}
