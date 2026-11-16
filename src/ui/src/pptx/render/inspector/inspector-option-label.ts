/**
 * The caption of a select field's choice, the one rule every binding's generic
 * renderer applies: a choice from the document (a slide title, a custom show
 * name) shows its own label, any other its translated key.
 *
 * @module render/inspector/inspector-option-label
 */
import type { InspectorFieldOption } from './inspector-field-model';

/** The caption a binding shows for `option`: its own label, else the translated key. */
export function inspectorOptionLabel(
	option: InspectorFieldOption,
	t: (key: string) => string,
): string {
	return option.label ?? t(option.labelKey);
}
