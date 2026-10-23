import { inspectorFieldPatch, inspectorTogglePatch, resolveInspectorSections } from 'ooxml-ui/pptx';
import type {
	InspectorFieldValue,
	InspectorFieldView,
	InspectorSectionDescriptor,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { WebCheckbox, WebSelect } from '../WebControls';
import { COL2, LBL, NUM, RNG, SEL } from './FillStrokeHelpers';
import { LiveColorInput } from './LiveColorInput';

export interface InspectorFieldsProps<T, P> {
	/** The section descriptors to draw (from `ooxml-ui/pptx`). */
	sections: readonly InspectorSectionDescriptor<T, P>[];
	/** The value the descriptors read. */
	target: T;
	/** Receives the patch a control change produces. */
	onPatch: (patch: P) => void;
	disabled?: boolean;
	/**
	 * For the colour fields: the value the descriptors read on any element, and
	 * how a field patch lands on it. With both, a drag through a colour picker
	 * previews and is put back leaf by leaf on the element it started on
	 * ({@link LiveColorInput}); without them every shade goes through `onPatch`.
	 */
	liveTarget?: (el: PptxElement) => T;
	livePatch?: (el: PptxElement, patch: P) => Partial<PptxElement>;
}

/**
 * The React binding's one generic field renderer: draws any list of shared
 * inspector section descriptors (a toggle per section, then its fields) as
 * cells of the surrounding two-column grid. Values, bounds, labels and the
 * patches all come from the descriptors; this component only maps them onto
 * JSX. A colour field shows every shade of a drag live but is one undo step
 * and one "Recent colours" entry when the pick settles ({@link LiveColorInput}).
 */
export function InspectorFields<T, P>({
	sections,
	target,
	onPatch,
	disabled = false,
	liveTarget,
	livePatch,
}: InspectorFieldsProps<T, P>): React.ReactElement {
	const { t } = useTranslation();
	const views = resolveInspectorSections(sections, target);

	const change = (sectionId: string, fieldId: string, raw: InspectorFieldValue): void => {
		const patch = inspectorFieldPatch(sections, sectionId, fieldId, target, raw);
		if (patch !== undefined) {
			onPatch(patch);
		}
	};

	const renderField = (sectionId: string, field: InspectorFieldView): React.ReactElement => {
		const label = t(field.labelKey);
		const cls = `flex flex-col gap-1 ${field.wide ? COL2 : ''}`;
		const key = `${sectionId}.${field.id}`;
		if (field.control === 'checkbox') {
			return (
				<label key={key} className={`flex items-center gap-2 ${field.wide ? COL2 : ''}`}>
					<WebCheckbox
						data-testid={field.testId}
						checked={Boolean(field.value)}
						disabled={disabled}
						onChange={(e) => change(sectionId, field.id, e.target.checked)}
						className='h-4 w-4'
					/>
					<span className={LBL}>{label}</span>
				</label>
			);
		}
		if (field.control === 'select') {
			return (
				<label key={key} className={cls}>
					<span className={LBL}>{label}</span>
					<WebSelect
						data-testid={field.testId}
						aria-label={label}
						value={String(field.value)}
						disabled={disabled}
						onChange={(e) => change(sectionId, field.id, e.target.value)}
						className={SEL}
					>
						{field.options?.map((o) => (
							<option key={o.value} value={o.value}>
								{t(o.labelKey)}
							</option>
						))}
					</WebSelect>
				</label>
			);
		}
		if (field.control === 'color') {
			return (
				<label key={key} className={cls}>
					<span className={LBL}>{label}</span>
					<LiveColorInput
						data-testid={field.testId}
						aria-label={label}
						value={String(field.value)}
						disabled={disabled}
						onApply={(hex) => change(sectionId, field.id, hex)}
						patch={
							liveTarget && livePatch
								? (el, hex) => {
										const fieldPatch = inspectorFieldPatch(
											sections,
											sectionId,
											field.id,
											liveTarget(el),
											hex,
										);
										return fieldPatch === undefined ? undefined : livePatch(el, fieldPatch);
									}
								: undefined
						}
						className={`h-8 ${SEL} px-1`}
					/>
				</label>
			);
		}
		return (
			<label key={key} className={cls}>
				<span className={LBL}>{label}</span>
				<input
					type={field.control}
					data-testid={field.testId}
					aria-label={label}
					min={field.min}
					max={field.max}
					step={field.step}
					value={Number(field.value)}
					disabled={disabled}
					onChange={(e) => change(sectionId, field.id, e.target.value)}
					className={field.control === 'range' ? RNG : NUM}
				/>
			</label>
		);
	};

	return (
		<>
			{views.map((section) => (
				<React.Fragment key={section.id}>
					{section.toggleable ? (
						<label className={`inline-flex items-center gap-2 text-foreground ${COL2}`}>
							<WebCheckbox
								data-testid={section.toggleTestId}
								checked={section.on}
								disabled={disabled}
								onChange={(e) => {
									const patch = inspectorTogglePatch(
										sections,
										section.id,
										target,
										e.target.checked,
									);
									if (patch !== undefined) {
										onPatch(patch);
									}
								}}
							/>
							{t(section.labelKey ?? '')}
						</label>
					) : section.labelKey ? (
						<span className={`${LBL} ${COL2}`}>{t(section.labelKey)}</span>
					) : null}
					{section.fields.map((field) => renderField(section.id, field))}
				</React.Fragment>
			))}
		</>
	);
}
