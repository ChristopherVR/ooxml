/** Builds and updates Draw groups, commands and their accessible availability reasons. */
export function createDrawRibbonControls(doc: Document, t: (key: string) => string) {
	const groups: HTMLElement[] = [];
	const commands: Array<{ control: HTMLButtonElement; key: string; reason: () => string }> = [];
	const labelGroups: Array<{ group: HTMLElement; key: string }> = [];
	const group = (id: string, key: string) => {
		const element = doc.createElement('pptx-ui-ribbon-group');
		element.dataset.ribbonGroup = `draw.${id}`;
		const row = doc.createElement('div');
		row.className = 'advanced-row';
		element.append(row);
		groups.push(element);
		labelGroups.push({ group: element, key });
		return row;
	};
	const button = (
		parent: HTMLElement,
		id: string,
		key: string,
		reason: () => string,
		action: () => void,
	) => {
		const control = doc.createElement('button');
		control.type = 'button';
		control.dataset.ribbonControl = `draw.${id}`;
		control.dataset.pptxCompact = '';
		control.addEventListener('click', () => {
			if (!reason()) action();
		});
		commands.push({ control, key, reason });
		parent.append(control);
		return control;
	};
	function sync(customText?: HTMLElement) {
		for (const { group: element, key } of labelGroups) element.setAttribute('label', t(key));
		for (const { control, key, reason } of commands) {
			const why = reason();
			if (control !== customText) control.textContent = t(key);
			control.setAttribute('aria-disabled', String(Boolean(why)));
			control.title = t(why || key);
			control.setAttribute('aria-description', why ? t(why) : '');
		}
	}
	return { groups, group, button, sync };
}
