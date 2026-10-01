// @vitest-environment node
import { OFFICE_UI_TAGS, registerOfficeUi, installOfficeUiTheme, defineButton } from './index.js';

describe('server-side import', () => {
	it('imports without a DOM and registering is a no-op', () => {
		expect(typeof window).toBe('undefined');
		expect(() => registerOfficeUi()).not.toThrow();
		expect(() => defineButton()).not.toThrow();
		expect(installOfficeUiTheme()).toBe(false);
		expect(OFFICE_UI_TAGS.length).toBeGreaterThan(0);
	});
});
