import { expect, it, vi } from 'vitest';
import '@angular/compiler';
import { createApplication } from '@angular/platform-browser';
import { createEnvironmentInjector, provideZonelessChangeDetection } from '@angular/core';
import { ViewerBuildingBlocks } from './building-blocks';

it('Angular service commits the inline text draft through the controller', async () => {
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const injector = createEnvironmentInjector([ViewerBuildingBlocks], app.injector);
	const blocks = injector.get(ViewerBuildingBlocks);
	const spy = vi.spyOn(blocks.controller, 'commitTextEdit');
	expect(await blocks.commitTextEdit()).toMatchObject({ status: 'unchanged', reason: 'no-draft' });
	expect(spy).toHaveBeenCalledOnce();
	injector.destroy();
	app.destroy();
});
