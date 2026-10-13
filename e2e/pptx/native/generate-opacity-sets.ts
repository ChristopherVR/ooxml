import { createOpacitySetDeck } from '../../../src/core/pptx/__tests__/helpers/opacity-set-deck';

const directory = process.argv[2];
if (!directory) throw new Error('Pass an output directory');
for (const hold of [true, false]) {
	await Bun.write(
		`${directory}/${hold ? 'hold' : 'remove'}.pptx`,
		await createOpacitySetDeck(hold),
	);
}
