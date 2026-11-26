import type { LocalAddIn } from 'ooxml-core/automation/add-ins';

const installs = new WeakMap<Node, Map<string, LocalAddIn>>();
export const INSTALLED_ADD_IN_LIMIT = 32;

export function localAddIns(scope: Node): Map<string, LocalAddIn> {
	const installed = installs.get(scope) ?? new Map<string, LocalAddIn>();
	installs.set(scope, installed);
	return installed;
}
