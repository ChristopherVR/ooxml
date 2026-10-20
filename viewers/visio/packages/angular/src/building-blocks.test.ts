import { expect, it, vi } from 'vitest';
import '@angular/compiler';
import { createApplication } from '@angular/platform-browser';
import {
	createComponent,
	createEnvironmentInjector,
	provideZonelessChangeDetection,
} from '@angular/core';
import {
	ViewerBuildingBlocks,
	DiagramCanvasComponent,
	ToolbarComponent,
	PageNavigationComponent,
	InspectorComponent,
	ViewerBuildingBlocksComponent,
} from './building-blocks';
it('Angular shares its DI owner and signals through native canvas and optional controls', async () => {
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const injector = createEnvironmentInjector([ViewerBuildingBlocks], app.injector);
	const blocks = injector.get(ViewerBuildingBlocks);
	const hosts: HTMLElement[] = [];
	const components = [
		DiagramCanvasComponent,
		ToolbarComponent,
		PageNavigationComponent,
		InspectorComponent,
		ViewerBuildingBlocksComponent,
	].map((type) => {
		const host = document.createElement('div');
		hosts.push(host);
		document.body.append(host);
		const component = createComponent(type, { environmentInjector: injector, hostElement: host });
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		return component;
	});
	expect(components[0]!.instance.blocks).toBe(blocks);
	expect(hosts[0]!.shadowRoot!.querySelector('[part="canvas"]')).not.toBeNull();
	blocks.toolbarProps().onZoomChange(2);
	expect(blocks.toolbarProps().zoom).toBe(2);
	blocks.update({ canEdit: false });
	expect(blocks.mode()).toBe('view');
	blocks.toolbarProps().onZoomChange(3);
	expect(blocks.state().zoom).toBe(3);
	const event = vi.fn();
	blocks.documentError.subscribe(event);
	const cause = new Error('blob read failed');
	const source = new Blob();
	source.arrayBuffer = async () => {
		throw cause;
	};
	blocks.update({ content: source });
	expect(blocks.loading()).toBe(true);
	await vi.waitFor(() => expect(blocks.error()).toBe(cause));
	expect(event).toHaveBeenCalledWith(cause);
	const destroy = vi.spyOn(blocks.controller, 'destroy');
	components.forEach((component) => {
		app.detachView(component.hostView);
		component.destroy();
	});
	injector.destroy();
	app.destroy();
	expect(destroy).toHaveBeenCalledOnce();
	hosts.forEach((host) => host.remove());
});

it('Angular input/output owner initializes callbacks even when no inputs are bound', async () => {
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const host = document.createElement('div');
	document.body.append(host);
	const component = createComponent(ViewerBuildingBlocksComponent, {
		environmentInjector: app.injector,
		hostElement: host,
	});
	const changed = vi.fn();
	component.instance.documentError.subscribe(changed);
	app.attachView(component.hostView);
	component.changeDetectorRef.detectChanges();
	const controller = component.instance.blocks.controller;
	const cause = new Error('host-owned blank load failure');
	await expect(
		controller.loadSource(async () => {
			throw cause;
		}),
	).rejects.toBe(cause);
	expect(changed).toHaveBeenCalledExactlyOnceWith(cause);
	app.detachView(component.hostView);
	component.destroy();
	app.destroy();
	host.remove();
});
