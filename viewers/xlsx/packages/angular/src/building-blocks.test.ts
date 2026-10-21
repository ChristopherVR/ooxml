// @vitest-environment jsdom
import '@angular/compiler';
import { expect, it } from 'vitest';
import { nativeBuildingBlockContract } from '../../../scripts/native-building-block-contract';
it('angular composes native controls with reactive inputs and engine lifecycle', async () => {
	expect.hasAssertions();
	await nativeBuildingBlockContract('angular');
}, 30000);

import { createApplication } from '@angular/platform-browser';
import {
	createEnvironmentInjector,
	runInInjectionContext,
	inject,
	provideZonelessChangeDetection,
} from '@angular/core';
import { ViewerBuildingBlocks } from './building-blocks';
it('forwards native service events before any options effect runs', async () => {
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const injector = createEnvironmentInjector([ViewerBuildingBlocks], app.injector);
	const blocks = runInInjectionContext(injector, () => inject(ViewerBuildingBlocks));
	const failures: Error[] = [];
	let workbooks = 0;
	blocks.workbookError.subscribe((error) => failures.push(error));
	blocks.workbookChange.subscribe(() => workbooks++);
	const host = document.createElement('div');
	document.body.append(host);
	blocks.canvasProps.mount(host);
	try {
		blocks.handle.newWorkbook();
		expect(workbooks).toBeGreaterThan(0);
		await expect(blocks.handle.load(new Uint8Array([0x50, 0x4b, 3, 4]))).rejects.toThrow(/./);
		expect(failures).toHaveLength(1);
	} finally {
		injector.destroy();
		app.destroy();
		host.remove();
	}
});
