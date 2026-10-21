// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { nativeBuildingBlockContract } from '../../../scripts/native-building-block-contract';
it('svelte composes native controls with reactive inputs and engine lifecycle', async () => {
	expect.hasAssertions();
	await nativeBuildingBlockContract('svelte');
}, 30000);
