import { inspectorFieldPatch, inspectorTogglePatch, resolveInspectorSections } from 'ooxml-ui/pptx';
import type {
	InspectorFieldValue,
	InspectorFieldView,
	InspectorSectionDescriptor,
} from 'ooxml-ui/pptx';
import React from 'react';
import { useTranslation } from 'react-i18next';

import { WebCheckbox, WebSelect } from '../WebControls';
import { COL2, LBL, NUM, RNG, SEL } from './FillStrokeHelpers';
import { useRecentColors } from './RecentColorsContext';

export interface InspectorFieldsProps<T, P> {
	/** The section descriptors to draw (from `ooxml-ui/pptx`). */
	sections: readonly InspectorSectionDescriptor<T, P>[];
	/** The value the descriptors read. */
	target: T;
	/** Receives the patch a control change produces. */
	onPatch: (patch: P) => void;
	disabled?: boolean;
}

/**
 * The React binding's one generic field renderer: draws any list of shared
 * inspector section descriptors (a toggle per section, then its fields) as
 * cells of the surrounding two-column grid. Values, bounds, labels and the
 * patches all come from the descriptors; this component only maps them onto
 * JSX. A colour commit also pushes into the "Recent colours" list.
 */
export function InspectorFields<T, P>({
	sections,
	target,
	onPatch,
	disabled = false,
}: InspectorFieldsProps<T, P>): React.ReactElement {
	const { t } = useTranslation();
	const { pushColor } = useRecentColors();
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
					<input
						type='color'
						data-testid={field.testId}
						aria-label={label}
						value={String(field.value)}
						disabled={disabled}
						onChange={(e) => {
							change(sectionId, field.id, e.target.value);
							pushColor(e.target.value);
						}}
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
							{t(section.labelKey)}
						</label>
					) : (
						<span className={`${LBL} ${COL2}`}>{t(section.labelKey)}</span>
					)}
					{section.fields.map((field) => renderField(section.id, field))}
				</React.Fragment>
			))}
		</>
	);
}
