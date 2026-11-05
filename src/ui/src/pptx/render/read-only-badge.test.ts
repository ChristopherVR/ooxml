import { describe, expect, it } from 'vitest';

import { resolveReadOnlyBadge } from './read-only-badge';

describe('resolveReadOnlyBadge', () => {
	it('shows for a read-only deck and hides when editing is allowed', () => {
		expect(resolveReadOnlyBadge({ editable: false }).visible).toBe(true);
		expect(resolveReadOnlyBadge({ editable: true }).visible).toBe(false);
	});

	it('yields to Enable Editing in Protected View', () => {
		expect(resolveReadOnlyBadge({ editable: false, enableEditingAvailable: true }).visible).toBe(
			false,
		);
	});

	it('names the shared text key and a status role', () => {
		const view = resolveReadOnlyBadge({ editable: false });
		expect(view.textKey).toBe('pptx.toolbar.readOnly');
		expect(view.role).toBe('status');
	});
});
