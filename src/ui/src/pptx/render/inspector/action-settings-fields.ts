/**
 * The inspector's Action Settings (PowerPoint's Action Settings dialog: what a
 * click on, or a pointer over, an element does), as field descriptors every
 * binding draws with its one generic renderer, once per trigger.
 *
 * The target is the trigger's context: the action the element carries, the
 * type the user has picked but not finished (`pendingType`), and the deck's
 * slides and custom shows, which the slide and show choices are resolved from
 * at render time (`optionsOf`). A change is an {@link ActionSettingsChange}:
 * the type to hold as pending, and the action to commit once it can carry its
 * target (`canCommitActionType`): "Go to URL", "Go to Slide" and "Custom Show"
 * are held until their target is filled in, so an unfinished pick never stamps
 * an empty action onto the element.
 *
 * @module render/inspector/action-settings-fields
 */
import type { ElementAction, ElementActionType, PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { pptxActionToElementAction } from 'ooxml-core/pptx';

import {
	canCommitActionType,
	ELEMENT_ACTION_TYPE_OPTIONS,
	resolveActionType,
} from '../element-action-options';
import { resolveSlideTitle } from '../field-context';
import type {
	InspectorFieldDescriptor,
	InspectorFieldOption,
	InspectorSectionDescriptor,
} from './inspector-field-model';

/** One trigger's Action Settings context. */
export interface ActionSettingsTarget {
	trigger: 'click' | 'hover';
	/** The action the element carries for this trigger, if any. */
	action: ElementAction | undefined;
	/** The type picked but not yet committed (it still needs its target). */
	pendingType?: ElementActionType;
	/** The deck's slides in order, by title (`undefined` for an untitled slide). */
	slideTitles: readonly (string | undefined)[];
	customShows: readonly { id: string; name: string }[];
}

/** What an Action Settings control did. */
export interface ActionSettingsChange {
	/** The type to hold as picked. */
	pendingType: ElementActionType;
	/** The action to write to the element, once it can carry its target. */
	commit?: ElementAction;
}

export type ActionSettingsSection = InspectorSectionDescriptor<
	ActionSettingsTarget,
	ActionSettingsChange
>;
type ActionSettingsField = InspectorFieldDescriptor<ActionSettingsTarget, ActionSettingsChange>;

/** The type the trigger shows: the pending pick, else the committed one. */
export function actionSettingsType(target: ActionSettingsTarget): ElementActionType {
	return resolveActionType(target.pendingType, target.action?.type);
}

/** The change for `type` with the trigger's current targets, plus `edit`. */
function change(
	target: ActionSettingsTarget,
	type: ElementActionType,
	edit: Partial<ElementAction> = {},
): ActionSettingsChange {
	const action: ElementAction = {
		trigger: target.trigger,
		type,
		url: target.action?.url,
		slideIndex: target.action?.slideIndex,
		customShowId: target.action?.customShowId,
		returnAfter: target.action?.returnAfter,
		...edit,
	};
	return canCommitActionType(type, action)
		? { pendingType: type, commit: action }
		: { pendingType: type };
}

const typeIs =
	(...types: ElementActionType[]) =>
	(target: ActionSettingsTarget): boolean =>
		types.includes(actionSettingsType(target));

function urlField(type: ElementActionType, labelKey: string): ActionSettingsField {
	return {
		id: `${type}Url`,
		kind: 'text',
		labelKey,
		wide: true,
		testId: `action-${type}-url`,
		visibleWhen: typeIs(type),
		read: (target) => target.action?.url ?? '',
		write: (target, value) => change(target, type, { url: String(value) }),
	};
}

/** The slide choices: every slide by number and title, after a "pick one" entry. */
function slideOptions(target: ActionSettingsTarget): InspectorFieldOption[] {
	const options: InspectorFieldOption[] = [];
	if (target.action?.slideIndex === undefined) {
		options.push({ value: '', labelKey: 'pptx.action.slideNumberPlaceholder' });
	}
	target.slideTitles.forEach((title, index) => {
		options.push({
			value: String(index),
			labelKey: '',
			label: title ? `${index + 1}. ${title}` : String(index + 1),
		});
	});
	return options;
}

/** The custom show choices, after a "pick one" entry. */
function customShowOptions(target: ActionSettingsTarget): InspectorFieldOption[] {
	return [
		{ value: '', labelKey: 'pptx.hyperlink.customShowLabel' },
		...target.customShows.map((show) => ({ value: show.id, labelKey: '', label: show.name })),
	];
}

/** One trigger's fields (untitled: each binding heads it On Click or On Hover). */
export const ACTION_SETTINGS_SECTION: ActionSettingsSection = {
	id: 'actionSettings',
	fields: [
		{
			id: 'type',
			kind: 'select',
			labelKey: 'pptx.action.title',
			options: ELEMENT_ACTION_TYPE_OPTIONS,
			wide: true,
			testId: 'action-type',
			read: actionSettingsType,
			write: (target, value) => change(target, String(value) as ElementActionType),
		},
		urlField('url', 'pptx.action.gotoUrl'),
		urlField('openFile', 'pptx.hyperlink.actionOpenFile'),
		urlField('openPresentation', 'pptx.hyperlink.actionOpenPresentation'),
		urlField('runProgram', 'pptx.hyperlink.actionRunProgram'),
		{
			id: 'slide',
			kind: 'select',
			labelKey: 'pptx.action.gotoSlide',
			wide: true,
			testId: 'action-slide',
			visibleWhen: typeIs('slide'),
			optionsOf: slideOptions,
			read: (target) =>
				target.action?.slideIndex === undefined ? '' : String(target.action.slideIndex),
			write: (target, value) => change(target, 'slide', { slideIndex: Number(value) }),
		},
		{
			id: 'customShow',
			kind: 'select',
			labelKey: 'pptx.hyperlink.customShowLabel',
			wide: true,
			testId: 'pptx-action-custom-show',
			visibleWhen: typeIs('customShow'),
			optionsOf: customShowOptions,
			read: (target) => target.action?.customShowId ?? '',
			write: (target, value) => change(target, 'customShow', { customShowId: String(value) }),
		},
		{
			id: 'returnAfter',
			kind: 'toggle',
			labelKey: 'pptx.hyperlink.customShowReturn',
			testId: 'pptx-action-custom-show-return',
			visibleWhen: typeIs('customShow'),
			read: (target) => target.action?.returnAfter === true,
			write: (target, value) => change(target, 'customShow', { returnAfter: value === true }),
		},
	],
};

/** One trigger's fields as a section list, what a generic renderer takes. */
export const ACTION_SETTINGS_SECTIONS: readonly ActionSettingsSection[] = [ACTION_SETTINGS_SECTION];

/**
 * The Action Settings context of `element`'s `trigger`: its action (an action
 * that parses back as `none` still lends its URL or slide, so a half-made
 * action keeps its target), the pending pick and the deck's choices.
 */
export function actionSettingsTargetOf(
	element: PptxElement,
	trigger: 'click' | 'hover',
	pendingType: ElementActionType | undefined,
	/** The deck's slides, or only their count when a binding has no more. */
	slides: readonly PptxSlide[] | number,
	customShows: readonly { id: string; name: string }[],
): ActionSettingsTarget {
	const raw = trigger === 'click' ? element.actionClick : element.actionHover;
	const parsed = raw ? pptxActionToElementAction(raw, trigger) : undefined;
	const action: ElementAction | undefined = parsed
		? {
				...parsed,
				url: parsed.url ?? raw?.url,
				slideIndex: parsed.slideIndex ?? raw?.targetSlideIndex,
			}
		: undefined;
	const slideTitles =
		typeof slides === 'number'
			? Array.from({ length: slides }, (): string | undefined => undefined)
			: slides.map((slide) => resolveSlideTitle(slide));
	const target: ActionSettingsTarget = { trigger, action, slideTitles, customShows };
	if (pendingType !== undefined) {
		target.pendingType = pendingType;
	}
	return target;
}
