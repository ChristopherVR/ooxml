// Test double for `validator.generated.mjs` (see fail-closed.test.mjs): the real validator, with
// the two entry points the server calls able to fail on request. Not published.
import {
	applyUpdateReverts as realApplyUpdateReverts,
	createUpdateChecker as realCreateUpdateChecker,
} from './validator.generated.mjs';

export * from './validator.generated.mjs';

export const faults = { revert: false, validate: false, revertCalls: 0 };

export function applyUpdateReverts(doc, reverts) {
	faults.revertCalls++;
	if (faults.revert) throw new Error('injected revert failure');
	return realApplyUpdateReverts(doc, reverts);
}

export function createUpdateChecker(doc) {
	const checker = realCreateUpdateChecker(doc);
	return {
		...checker,
		validate(...args) {
			if (faults.validate) throw new Error('injected validator failure');
			return checker.validate(...args);
		},
	};
}
