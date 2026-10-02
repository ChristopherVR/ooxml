import { beforeEach, expect, it, vi } from 'vitest';
import { current, emit, reset } from './mock-binding.js';
import '@angular/compiler';
import { createApplication } from '@angular/platform-browser';
import {
	createComponent,
	provideZonelessChangeDetection,
	reflectComponentType,
} from '@angular/core';
import { VisioViewerComponent } from '../src/angular.js';
import { propertyKeys } from '../src/common.js';
beforeEach(reset);
it('Angular mounts with real native lifecycle and zoneless input/output updates', async () => {
	expect(
		reflectComponentType(VisioViewerComponent)
			?.inputs.map((input) => input.propName)
			.sort(),
	).toEqual([...propertyKeys, 'events'].sort());
	const host = document.createElement('div');
	document.body.append(host);
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const component = createComponent(VisioViewerComponent, {
		environmentInjector: app.injector,
		hostElement: host,
	});
	const first = vi.fn();
	const latest = vi.fn();
	const native = vi.fn();
	component.setInput('document', null);
	component.setInput('pageIndex', 0);
	component.setInput('zoom', 2);
	component.setInput('showToolbar', false);
	component.setInput('events', { 'zoom-change': first });
	component.instance.zoomChange.subscribe(native);
	app.attachView(component.hostView);
	component.changeDetectorRef.detectChanges();
	const live = current();
	const retained = component.instance;
	expect(live.options).toMatchObject({ document: null, pageIndex: 0, zoom: 2, showToolbar: false });
	component.setInput('zoom', 3);
	component.setInput('showToolbar', true);
	component.setInput('events', { 'zoom-change': latest });
	component.changeDetectorRef.detectChanges();
	emit('zoom-change', 3);
	expect(current()).toBe(live);
	expect(first).not.toHaveBeenCalled();
	expect(latest).toHaveBeenCalledWith(3);
	expect(native).toHaveBeenCalledWith(3);
	component.setInput('events', undefined);
	component.changeDetectorRef.detectChanges();
	emit('zoom-change', 4);
	expect(latest).toHaveBeenCalledOnce();
	app.detachView(component.hostView);
	component.destroy();
	app.destroy();
	expect(live.destroy).toHaveBeenCalledOnce();
	expect(() => retained.fit()).toThrow('not mounted');
	host.remove();
});
