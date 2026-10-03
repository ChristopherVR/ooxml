import type { Workbook } from '@christophervr/xlsx-core';
import { heading, labelled, primary, type PageContext } from './parts';

/** Edits known custom-property values while retaining unknown variants verbatim. */
export function customProperties(page: PageContext, workbook: Workbook): HTMLElement {
	const { doc, host, t } = page;
	const section = doc.createElement('section');
	const render = () => {
		section.replaceChildren(heading(page, t('Custom properties'), 'h3'));
		const properties = workbook.properties.custom ?? [];
		for (const [index, property] of properties.entries()) {
			const input = doc.createElement('input');
			input.type =
				property.type === 'bool'
					? 'checkbox'
					: property.type === 'i4' || property.type === 'r8'
						? 'number'
						: 'text';
			input.disabled = host.ctx.readOnly() || property.type === 'raw';
			if (property.type === 'bool') input.checked = property.value;
			else input.value = property.type === 'raw' ? property.xml : String(property.value);
			input.addEventListener('change', () => {
				if (host.ctx.readOnly() || property.type === 'raw' || host.ctx.workbook() !== workbook)
					return;
				let updated = property;
				if (property.type === 'bool') updated = { ...property, value: input.checked };
				else if (property.type === 'i4' || property.type === 'r8') {
					const value = Number(input.value);
					if (
						!input.value.trim() ||
						!Number.isFinite(value) ||
						(property.type === 'i4' &&
							(!Number.isInteger(value) || value < -2147483648 || value > 2147483647))
					) {
						input.value = String(property.value);
						return;
					}
					updated = { ...property, value };
				} else if (property.type === 'lpwstr' || property.type === 'filetime') {
					if (property.type === 'filetime' && !Number.isFinite(Date.parse(input.value))) {
						input.value = property.value;
						return;
					}
					updated = { ...property, value: input.value };
				}
				host.setProperty(
					'custom',
					(workbook.properties.custom ?? []).map((p, i) => (i === index ? updated : p)),
				);
			});
			const remove = primary(page, t('Remove'), () => {
				if (host.ctx.readOnly() || host.ctx.workbook() !== workbook) return;
				host.setProperty(
					'custom',
					(workbook.properties.custom ?? []).filter((_, i) => i !== index),
				);
				render();
			});
			remove.disabled = host.ctx.readOnly();
			section.append(labelled(page, property.name, input), remove);
		}
		const name = doc.createElement('input');
		const value = doc.createElement('input');
		name.disabled = value.disabled = host.ctx.readOnly();
		const add = primary(page, t('Add text property'), () => {
			if (host.ctx.readOnly() || host.ctx.workbook() !== workbook) return;
			const key = name.value.trim();
			const current = workbook.properties.custom ?? [];
			if (!key || current.some((p) => p.name.toLowerCase() === key.toLowerCase())) return;
			host.setProperty('custom', [...current, { name: key, type: 'lpwstr', value: value.value }]);
			render();
		});
		add.disabled = host.ctx.readOnly();
		section.append(
			labelled(page, t('Property name'), name),
			labelled(page, t('Property value'), value),
			add,
		);
	};
	render();
	return section;
}
