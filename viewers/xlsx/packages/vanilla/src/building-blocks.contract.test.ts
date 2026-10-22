// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { nativeBuildingBlockContract } from '../../../scripts/native-building-block-contract';
it('vanilla composes native controls with reactive inputs and engine lifecycle', async () => {
	expect.hasAssertions();
	await nativeBuildingBlockContract('vanilla');
}, 30000);
