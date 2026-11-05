import { describe, expect, it } from 'vitest';

import { buildEquationBlockStyle, EQUATION_FONT_FAMILY } from './equation-box-style';

describe('buildEquationBlockStyle', () => {
	it('keeps the anchor and insets of the text block and only centres the cross axis', () => {
		const style = buildEquationBlockStyle({
			display: 'flex',
			flexDirection: 'column',
			justifyContent: 'flex-start',
			paddingTop: '3px',
		});
		expect(style.flexDirection).toBe('column');
		expect(style.justifyContent).toBe('flex-start');
		expect(style.paddingTop).toBe('3px');
		expect(style.alignItems).toBe('center');
		expect(style.flexWrap).toBe('nowrap');
		expect(style.fontFamily).toBe(EQUATION_FONT_FAMILY);
	});
});
