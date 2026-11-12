/**
 * ribbon-display-menu.component.ts: the Ribbon Display Options menu, drawn from the shared
 * controller's state. A zero-height anchor sits right after the tab row, so the menu hangs under
 * the tab row and is not clipped by the scrolling tab strip.
 */
import { ChangeDetectionStrategy, Component, computed, input } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { RIBBON_DISPLAY_OPTIONS_LABEL_KEY } from 'ooxml-ui/pptx';
import type { RibbonDisplayController, RibbonDisplayState } from 'ooxml-ui/pptx';

@Component({
	selector: 'pptx-ribbon-display-menu',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe],
	template: `
		<div class="pointer-events-none relative h-0 w-full">
			@if (open()) {
				<div
					role="menu"
					data-pptx-chrome="ribbon-display-menu"
					class="pointer-events-auto absolute right-1.5 top-0.5 z-50 min-w-[220px] rounded-md border border-border bg-card py-1 shadow-lg"
					[attr.aria-label]="labelKey | translate"
				>
					@for (item of menu(); track item.mode) {
						<button
							type="button"
							role="menuitemradio"
							class="relative flex w-full items-center whitespace-nowrap py-1.5 pl-7 pr-3 text-left text-[12px] text-foreground transition-colors hover:bg-accent/30"
							[attr.aria-checked]="item.checked"
							[attr.data-ribbon-display]="item.mode"
							(click)="controller()?.choose(item.mode)"
						>
							@if (item.checked) {
								<span aria-hidden="true" class="absolute left-2.5 text-[11px] text-primary"
									>&#10003;</span
								>
							}
							{{ item.labelKey | translate }}
						</button>
					}
				</div>
			}
		</div>
	`,
})
export class RibbonDisplayMenuComponent {
	/** The shared controller's state; nothing is drawn without it. */
	readonly state = input<RibbonDisplayState | null>(null);
	/** The shared controller the menu items call. */
	readonly controller = input<RibbonDisplayController | null>(null);

	protected readonly labelKey = RIBBON_DISPLAY_OPTIONS_LABEL_KEY;
	protected readonly open = computed(() => this.state()?.menuOpen ?? false);
	protected readonly menu = computed(() => this.state()?.menu ?? []);
}
