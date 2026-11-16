import {
	ChangeDetectionStrategy,
	Component,
	computed,
	inject,
	input,
	output,
	signal,
} from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { ACTION_SETTINGS_SECTIONS, actionSettingsTargetOf, resolveActionType } from 'ooxml-ui/pptx';
import type { ActionSettingsChange, ActionSettingsTarget } from 'ooxml-ui/pptx';
import type { ElementActionType, PptxCustomShow, PptxElement } from 'pptx-viewer-core';
import { elementActionToPptxAction } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';
import { LoadContentService } from './load-content.service';

type Trigger = 'click' | 'hover';

/** A type the user picked, remembered against the element it was picked on. */
export interface PendingActionType {
	elementId: string;
	types: Partial<Record<Trigger, ElementActionType>>;
}

/** Nothing picked yet. */
export const NO_PENDING_ACTION_TYPE: PendingActionType = { elementId: '', types: {} };

/**
 * Record the type the user just picked for one trigger.
 *
 * A pick belongs to the element it was made on, so a pick for a different
 * element replaces the whole record instead of merging: otherwise selecting
 * another shape would show it a half-made choice it never had.
 */
export function withPendingActionType(
	previous: PendingActionType,
	elementId: string,
	trigger: Trigger,
	type: ElementActionType,
): PendingActionType {
	const types = previous.elementId === elementId ? previous.types : {};
	return { elementId, types: { ...types, [trigger]: type } };
}

/**
 * The action type a trigger's controls should render, which is also the
 * predicate deciding whether the URL / slide input exists.
 *
 * Exported (and pure) so it can be unit-tested without a TestBed, matching the
 * rest of this package; see `inspector-panel.component.ts` for the convention.
 */
export function displayedActionType(
	pending: PendingActionType,
	elementId: string,
	trigger: Trigger,
	committedType: ElementActionType | undefined,
): ElementActionType {
	const picked = pending.elementId === elementId ? pending.types[trigger] : undefined;
	return resolveActionType(picked, committedType);
}

@Component({
	selector: 'pptx-action-settings-panel',
	standalone: true,
	imports: [TranslatePipe, InspectorFieldsComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<section class="pptx-ng-action" data-pptx-action-settings>
			<h3 class="pptx-ng-action__heading">{{ 'pptx.action.title' | translate }}</h3>
			@for (trigger of triggers; track trigger) {
				<div class="pptx-ng-action__trigger" [attr.data-pptx-action-trigger]="trigger">
					<span class="pptx-ng-action__label">
						{{ (trigger === 'click' ? 'pptx.action.onClick' : 'pptx.action.onHover') | translate }}
					</span>
					<pptx-inspector-fields
						[sections]="sections"
						[target]="targetFor(trigger)"
						(patch)="onChange(trigger, $event)"
					/>
				</div>
			}
		</section>
	`,
	styles: `
		.pptx-ng-action {
			display: grid;
			gap: 8px;
			font-size: 11px;
		}
		.pptx-ng-action__heading {
			margin: 0;
			font-size: 11px;
			font-weight: 600;
		}
		.pptx-ng-action__label {
			color: var(--pptx-inspector-muted, #aaa);
		}
	`,
})
export class ActionSettingsPanelComponent {
	readonly element = input.required<PptxElement>();
	readonly slideCount = input(0);
	readonly patch = output<Partial<PptxElement>>();

	protected readonly triggers = ['click', 'hover'] as const;
	/** One trigger's shared Action Settings descriptors (`ACTION_SETTINGS_SECTIONS`). */
	protected readonly sections = ACTION_SETTINGS_SECTIONS;
	/**
	 * The type the user just picked, per trigger, tagged with its element id so
	 * moving the inspector to another shape does not carry a half-made choice
	 * across.
	 */
	private readonly pending = signal<PendingActionType>(NO_PENDING_ACTION_TYPE);

	/** `data.customShows`, for the `customShow` target choices. */
	private readonly loader = inject(LoadContentService, { optional: true });
	protected readonly customShows = computed<readonly PptxCustomShow[]>(
		() => this.loader?.customShows() ?? [],
	);

	/** One trigger's context: its action, the pending pick and the deck's choices. */
	protected targetFor(trigger: Trigger): ActionSettingsTarget {
		const pending = this.pending();
		const element = this.element();
		const picked = pending.elementId === element.id ? pending.types[trigger] : undefined;
		return actionSettingsTargetOf(
			element,
			trigger,
			picked,
			this.slideCount(),
			this.customShows().map((show) => ({ id: show.id, name: show.name })),
		);
	}

	protected onChange(trigger: Trigger, change: ActionSettingsChange): void {
		this.pending.set(
			withPendingActionType(this.pending(), this.element().id, trigger, change.pendingType),
		);
		if (change.commit) {
			const action = elementActionToPptxAction(change.commit);
			this.patch.emit(
				(trigger === 'click'
					? { actionClick: action }
					: { actionHover: action }) as Partial<PptxElement>,
			);
		}
	}
}
