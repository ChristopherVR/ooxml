import { it } from 'vitest';
import { writeFileSync } from 'node:fs';
import { evaluatePresetShape } from './preset-shape-evaluator';
it('dump', () => {
	const r = evaluatePresetShape('actionButtonHelp', 1000, 1000, undefined)!;
	writeFileSync(process.env.SWOUT!, JSON.stringify(r.paths));
});
