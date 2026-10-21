// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { nativeBuildingBlockContract } from '../../../scripts/native-building-block-contract';
it('vue composes native controls with reactive inputs and engine lifecycle', async () => {
	expect.hasAssertions();
	await nativeBuildingBlockContract('vue');
}, 30000);

import { createApp, h, shallowRef, nextTick, effectScope } from 'vue';
import { SpreadsheetCanvas, useViewerBuildingBlocks } from './building-blocks';
import { createBuildingBlockSession } from '../../bindings/src/building-blocks';
import { createWorkbook } from '../../core/src';
it('cleans effect scopes and transfers a canvas to a new mount callback', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const first = createBuildingBlockSession(() => {});
	const second = createBuildingBlockSession(() => {});
	first.update({ workbook: createWorkbook() });
	second.update({ workbook: createWorkbook() });
	const mount = shallowRef(first.mount);
	const app = createApp({ setup: () => () => h(SpreadsheetCanvas, { mount: mount.value }) });
	app.mount(host);
	const old = host.querySelector('xlsx-editor')!;
	mount.value = second.mount;
	await nextTick();
	expect(old.isConnected).toBe(false);
	expect(host.querySelectorAll('xlsx-editor')).toHaveLength(1);
	expect(await first.handle.saveBytes().catch(() => null)).toBeNull();
	app.unmount();
	expect(host.querySelector('xlsx-editor')).toBeNull();
	const scope = effectScope();
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	scope.run(() => {
		blocks = useViewerBuildingBlocks(() => ({ workbook: createWorkbook() }));
	});
	blocks.canvasProps.mount(host);
	expect(host.querySelector('xlsx-editor')).not.toBeNull();
	scope.stop();
	expect(host.querySelector('xlsx-editor')).toBeNull();
	host.remove();
});
