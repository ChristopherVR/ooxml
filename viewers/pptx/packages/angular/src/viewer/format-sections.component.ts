/**
 * format-sections.component.ts: the inspector's Properties tab as a
 * selection-aware Format pane. Renders the shared section list for the
 * selection (`resolveFormatSections`) as `pptx-ui-format-section` disclosures
 * and projects, into each, the host's `<ng-template pptxFormatSection="id">`
 * blocks for that id (the existing panels, in their own template order).
 *
 * Open state lives in the viewer's Format pane controller
 * (`ViewerInspectorPanelService.formatPane`), so a closed section stays closed
 * for the next element of the same category; a launcher request
 * (`requestFormatPane`) is taken after render, opening, revealing and focusing
 * its section. The pane scrolls to the top when the category changes.
 */
import { NgTemplateOutlet } from '@angular/common';
import {
	ChangeDetectionStrategy,
	Component,
	computed,
	ContentChildren,
	CUSTOM_ELEMENTS_SCHEMA,
	Directive,
	effect,
	ElementRef,
	inject,
	Input,
	input,
	QueryList,
	TemplateRef,
	untracked,
} from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import {
	createFormatPaneController,
	formatSectionOpen,
	isFormatSectionMixed,
	resolveFormatCategory,
	resolveFormatSections,
	revealFormatSection,
} from 'ooxml-ui/pptx';
import type { FormatSectionId, FormatSectionView } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { controllerSignal } from './controller-signal';
import { ViewerInspectorPanelService } from './viewer-inspector-panel.service';

/** One section body: `<ng-template pptxFormatSection="fill">...</ng-template>`. */
@Directive({ selector: 'ng-template[pptxFormatSection]', standalone: true })
export class FormatSectionTemplateDirective {
	// A decorator input and query (not signal ones), so the package's plain JIT tests see them.
	@Input({ required: true, alias: 'pptxFormatSection' }) section!: FormatSectionId;
	readonly template = inject<TemplateRef<{ $implicit: FormatSectionView }>>(TemplateRef);
}

@Component({
	selector: 'pptx-format-sections',
	standalone: true,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [NgTemplateOutlet, TranslatePipe],
	template: `
		<div
			class="format-sections"
			role="group"
			[attr.data-pptx-format-sections]="category()"
			[attr.aria-label]="'pptx.formatPane.sections' | translate"
		>
			@for (view of views(); track view.id) {
				<pptx-ui-format-section
					[attr.data-pptx-format-section]="view.id"
					[label]="view.labelKey | translate"
					[open]="isOpen(view)"
					[mixed]="isMixed(view)"
					[mixedLabel]="'pptx.formatPane.mixed' | translate"
					(office-section-toggle)="onToggle(view, $event)"
				>
					@for (body of bodiesFor(view.id); track $index) {
						<ng-container *ngTemplateOutlet="body; context: { $implicit: view }" />
					}
				</pptx-ui-format-section>
			}
		</div>
	`,
	styles: `
		:host {
			display: block;
		}
		.format-sections {
			display: flex;
			flex-direction: column;
		}
	`,
})
export class FormatSectionsComponent {
	/** The selection (primary first); empty for the slide. */
	readonly elements = input<readonly PptxElement[]>([]);

	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private readonly inspector = inject(ViewerInspectorPanelService, { optional: true });
	/** The viewer's controller, or a private one when rendered on its own. */
	protected readonly pane = this.inspector?.formatPane ?? createFormatPaneController();
	private readonly state = controllerSignal(this.pane);
	@ContentChildren(FormatSectionTemplateDirective)
	private bodies?: QueryList<FormatSectionTemplateDirective>;

	protected readonly views = computed(() => resolveFormatSections(this.elements()));
	protected readonly category = computed(() => resolveFormatCategory(this.elements()));

	constructor() {
		effect(() => {
			const category = this.category();
			const views = this.views();
			// Re-run when a request arrives; take it without tracking the writes it makes.
			this.state().request;
			untracked(() => {
				const changed = this.pane.noteCategory(category);
				const reveal = this.pane.takeRequest(category, views);
				const root = this.host.nativeElement;
				if (reveal) {
					requestAnimationFrame(() => revealFormatSection(root, reveal.section, reveal));
				} else if (changed) {
					const scroller = root.closest<HTMLElement>('[data-pptx-chrome="inspector-body"]');
					if (scroller) {
						scroller.scrollTop = 0;
					}
				}
			});
		});
	}

	protected isOpen(view: FormatSectionView): boolean {
		return formatSectionOpen(this.state(), this.category(), view);
	}

	protected isMixed(view: FormatSectionView): boolean {
		return isFormatSectionMixed(this.elements(), view.id);
	}

	protected bodiesFor(id: FormatSectionId): TemplateRef<{ $implicit: FormatSectionView }>[] {
		return (this.bodies?.toArray() ?? [])
			.filter((body) => body.section === id)
			.map((body) => body.template);
	}

	protected onToggle(view: FormatSectionView, event: Event): void {
		const open = (event as CustomEvent<{ open: boolean }>).detail.open;
		this.pane.setOpen(this.category(), view.id, open);
	}
}
