/**
 * ribbon-reveal-strip.component.ts: the slim strip of Full-screen mode that brings the ribbon back
 * over the slide. The ribbon host draws it only in full-screen mode; a press calls `reveal()`.
 */
import { ChangeDetectionStrategy, Component, input } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { RIBBON_DISPLAY_REVEAL_LABEL_KEY } from 'ooxml-ui/pptx';
import type { RibbonDisplayController } from 'ooxml-ui/pptx';

@Component({
	selector: 'pptx-ribbon-reveal-strip',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe],
	template: `
		<button
			type="button"
			data-pptx-chrome="ribbon-reveal"
			class="block h-3.5 w-full shrink-0 border-b border-border bg-secondary/50 text-[11px] leading-none text-muted-foreground transition-colors hover:text-foreground"
			[attr.aria-label]="labelKey | translate"
			[title]="labelKey | translate"
			(click)="controller()?.reveal()"
		>
			···
		</button>
	`,
})
export class RibbonRevealStripComponent {
	/** The shared controller whose `reveal()` the strip calls. */
	readonly controller = input<RibbonDisplayController | null>(null);

	protected readonly labelKey = RIBBON_DISPLAY_REVEAL_LABEL_KEY;
}
