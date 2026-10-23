/**
 * The colour fields of a generic field renderer, live while the user drags
 * through their native pickers: one shared `createLiveColorInput` controller
 * per colour field, created on first use, so every binding's renderer only
 * routes its colour input's events to `colorField(sectionId, fieldId)`.
 *
 * Each shade is the field's descriptor patch on the renderer's target,
 * previewed for the element the target belongs to (`targetKey`); the settled
 * pick is the field's one `patch`. With the element's identity in the key, a
 * field tells a new selection from a re-render of its own drag, so a drag the
 * selection leaves is put back on its own element and never lands on the next.
 *
 * @module render/inspector/inspector-color-fields
 */
import type { LiveColorInput } from '../live-color-input';
import { createLiveColorInput } from '../live-color-input';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import type { InspectorSectionDescriptor } from './inspector-field-model';

/** What a renderer gives {@link createInspectorColorFields}; read afresh on every event. */
export interface InspectorColorFieldsOptions<T, P> {
	sections(): readonly InspectorSectionDescriptor<T, P>[];
	/** The value the descriptors read now. */
	target(): T;
	/** The identity of the element `target` belongs to (its id). */
	targetKey(): string;
	/**
	 * What the descriptors read on element `key` now (its `target`), or `undefined` when
	 * it is gone: what a drag the selection has left is put back on.
	 */
	lookup(key: string): T | undefined;
	/** A shade, or the drag put back, for element `target`: apply it without an undo step. */
	preview(patch: P, target: string): void;
	/** The settled pick for the current target: the field's normal patch (one undo step). */
	commit(patch: P): void;
	/** A settled pick for the recent colours. */
	remember(hex: string): void;
}

/** The live colour controller of each colour field, keyed by section and field. */
export function createInspectorColorFields<T, P>(
	options: InspectorColorFieldsOptions<T, P>,
): (sectionId: string, fieldId: string) => LiveColorInput {
	const controls = new Map<string, LiveColorInput>();
	return (sectionId, fieldId) => {
		const id = `${sectionId}.${fieldId}`;
		let control = controls.get(id);
		if (!control) {
			control = createLiveColorInput<T, P & object>({
				read: () => {
					const target = options.target();
					const field = resolveInspectorSections(options.sections(), target)
						.find((section) => section.id === sectionId)
						?.fields.find((candidate) => candidate.id === fieldId);
					return field && { key: options.targetKey(), color: String(field.value), value: target };
				},
				patch: (target, hex) =>
					inspectorFieldPatch(options.sections(), sectionId, fieldId, target, hex) as
						| (P & object)
						| undefined,
				lookup: (key) => options.lookup(key),
				preview: (patch, target) => options.preview(patch, target),
				commit: (patch) => options.commit(patch),
				remember: (hex) => options.remember(hex),
			});
			controls.set(id, control);
		}
		return control;
	};
}
