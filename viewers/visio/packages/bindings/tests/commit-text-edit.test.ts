import { expect, it } from 'vitest';
import { createBuildingBlockSession } from '../src/building-blocks';
import { createViewerBuildingBlocks } from '../src/vanilla-building-blocks';

it('vanilla blocks commit the inline text draft', async () => {
	const blocks = createViewerBuildingBlocks({ canEdit: true });
	try {
		expect(await blocks.commitTextEdit()).toMatchObject({
			status: 'unchanged',
			reason: 'no-draft',
		});
	} finally {
		blocks.destroy();
	}
});

it('the shared session commits the inline text draft and reports a destroyed owner', async () => {
	const session = createBuildingBlockSession({ canEdit: true });
	expect(await session.commitTextEdit()).toMatchObject({ status: 'unchanged' });
	session.destroy();
	expect(await session.commitTextEdit()).toMatchObject({
		status: 'cancelled',
		reason: 'destroyed',
	});
});
