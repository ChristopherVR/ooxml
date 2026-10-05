import type { VisioDiagnostic } from 'ooxml-core/visio';
export interface CompatibilityNote {
	message: string;
	count: number;
	severity: 'info' | 'warning';
}
/** Presentation-only grouping keeps repeated per-shape diagnostics readable without losing counts. */
export function compatibilityNotes(
	diagnostics: readonly VisioDiagnostic[],
	renderWarnings: readonly string[] = [],
): CompatibilityNote[] {
	const notes = new Map<string, CompatibilityNote>();
	for (const diagnostic of diagnostics) {
		const existing = notes.get(diagnostic.message);
		if (existing) {
			existing.count++;
			if (diagnostic.severity === 'warning') existing.severity = 'warning';
		} else
			notes.set(diagnostic.message, {
				message: diagnostic.message,
				count: 1,
				severity: diagnostic.severity,
			});
	}
	for (const message of renderWarnings) {
		const existing = notes.get(message);
		if (existing) {
			existing.count++;
			existing.severity = 'warning';
		} else notes.set(message, { message, count: 1, severity: 'warning' });
	}
	return [...notes.values()].sort(
		(a, b) => Number(a.severity === 'info') - Number(b.severity === 'info'),
	);
}
export function compatibilityText(note: CompatibilityNote): string {
	return note.message + (note.count > 1 ? ` (${note.count} occurrences)` : '');
}
