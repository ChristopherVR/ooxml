import { convertLatexToOmml, parseOmml } from '../../../math/index';
import type { SecurityCase } from '../case';

export const cases: readonly SecurityCase[] = [
	{
		name: 'a stray closing brace terminates malformed LaTeX parsing',
		limits: { ms: 1500, heapMb: 128 },
		run: () => {
			try {
				convertLatexToOmml('}');
			} catch (error) {
				return error instanceof Error ? error.message : 'unknown error';
			}
		},
		expected: 'Unexpected closing brace',
	},
	{
		name: 'an oversized Office Math run is refused before adaptation',
		limits: { ms: 1500, heapMb: 128 },
		run: () => {
			const text = 'x'.repeat(262_145);
			try {
				parseOmml(
					`<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>${text}</m:t></m:r></m:oMath>`,
				);
			} catch (error) {
				return error instanceof Error ? error.message : 'unknown error';
			}
		},
		expected: 'OMML exceeds the size limit',
	},
];
