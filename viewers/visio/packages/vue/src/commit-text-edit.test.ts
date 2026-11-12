import { expect, it, vi } from 'vitest';
import { createApp, defineComponent, h } from 'vue';
import { DiagramCanvas, useViewerBuildingBlocks } from './building-blocks';

it('Vue composable commits the inline text draft through the controller', async () => {
	const host = document.createElement('div');
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	const app = createApp(
		defineComponent({
			setup() {
				blocks = useViewerBuildingBlocks(() => ({ canEdit: true }));
				return () => h(DiagramCanvas, blocks.canvasProps);
			},
		}),
	);
	app.mount(host);
	const spy = vi.spyOn(blocks.inspectorProps.controller, 'commitTextEdit');
	expect(await blocks.commitTextEdit()).toMatchObject({ status: 'unchanged', reason: 'no-draft' });
	expect(spy).toHaveBeenCalledOnce();
	app.unmount();
});
