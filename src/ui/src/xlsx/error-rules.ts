// The reader's Error Checking rule toggles (File > Options > Formulas), application wide like
// Excel's custom lists: kept in the browser's localStorage when available and for the session
// otherwise. Defaults to Excel's default rules (`DEFAULT_ERROR_RULES`).
import { DEFAULT_ERROR_RULES, ERROR_RULES, type ErrorRule } from 'ooxml-core/xlsx';

/** The localStorage key for the enabled Error Checking rules (JSON array of `ErrorRule`). */
export const ERROR_RULES_KEY = 'ooxml.xlsx.errorCheckingRules';

const VALID_RULES = new Set<string>(ERROR_RULES);

let memory: ErrorRule[] | undefined;

function parseRules(raw: string): ErrorRule[] | undefined {
	try {
		const parsed: unknown = JSON.parse(raw);
		if (!Array.isArray(parsed)) return undefined;
		return parsed.filter(
			(item): item is ErrorRule => typeof item === 'string' && VALID_RULES.has(item),
		);
	} catch {
		return undefined;
	}
}

/** The saved Error Checking rules, or the session copy, or Excel's defaults. */
export function savedErrorRules(): ErrorRule[] {
	try {
		const raw = globalThis.localStorage?.getItem(ERROR_RULES_KEY);
		if (raw !== null && raw !== undefined) {
			const parsed = parseRules(raw);
			if (parsed) return parsed;
		}
	} catch {
		// Storage can be blocked; the session value applies.
	}
	return memory ? [...memory] : [...DEFAULT_ERROR_RULES];
}

/** Persists the enabled Error Checking rules. */
export function saveErrorRules(rules: readonly ErrorRule[]): void {
	const filtered = rules.filter((rule) => VALID_RULES.has(rule));
	memory = [...filtered];
	try {
		globalThis.localStorage?.setItem(ERROR_RULES_KEY, JSON.stringify(memory));
	} catch {
		// Storage can be blocked; the choice then lasts for the session.
	}
}

/** Forgets the saved rules (tests, and a reset). */
export function clearSavedErrorRules(): void {
	memory = undefined;
	try {
		globalThis.localStorage?.removeItem(ERROR_RULES_KEY);
	} catch {
		// Nothing to remove.
	}
}
