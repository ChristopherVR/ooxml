/**
 * viewer-region-host.component.ts: `<pptx-region-host>`, the one place the
 * Angular binding hosts a host-supplied region renderer.
 *
 * It owns a single `div` container and a shared {@link ViewerRegionHost}: the
 * host mounts the renderer, delivers context updates and tears it down, so the
 * semantics are identical in every binding. This component only forwards the
 * `renderer` and `context` inputs and destroys the host with the view.
 */
import {
	ChangeDetectionStrategy,
	Component,
	DestroyRef,
	effect,
	ElementRef,
	inject,
	input,
} from '@angular/core';
import { ViewerRegionHost } from 'ooxml-ui/pptx';
import type { ViewerRegionContext, ViewerRegionRenderer } from 'ooxml-ui/pptx';

@Component({
	selector: 'pptx-region-host',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { style: 'display: block; width: 100%; height: 100%; min-height: 0; min-width: 0' },
	template: '',
})
export class ViewerRegionHostComponent {
	/** The host renderer: a mount function, a tag name or `{ element, properties }`. */
	readonly renderer = input.required<ViewerRegionRenderer>();
	/** What the renderer receives; a new value is delivered as an update. */
	readonly context = input.required<ViewerRegionContext>();

	constructor() {
		const element = inject<ElementRef<HTMLElement>>(ElementRef).nativeElement;
		const container = element.ownerDocument.createElement('div');
		container.className = 'pptx-ng-region-host';
		container.style.cssText = 'width: 100%; height: 100%; min-height: 0; min-width: 0';
		element.append(container);
		const host = new ViewerRegionHost(container);
		effect(() => host.render(this.renderer(), this.context()));
		inject(DestroyRef).onDestroy(() => {
			host.destroy();
			container.remove();
		});
	}
}
