import { createElement, type MouseEvent, type ReactNode } from 'react';
import type { Control, ControlGroup } from '../../bindings/src/controls';

const keepSelection = (event: MouseEvent) => event.preventDefault();

function renderControl(control: Control): ReactNode {
	if (control.kind === 'text')
		return createElement('span', { key: control.id, role: 'status' }, control.text);
	if (control.kind === 'select')
		return createElement(
			'select',
			{
				key: control.id,
				'aria-label': control.label,
				disabled: control.disabled,
				value: control.value,
				onChange: (event: { target: { value: string } }) => control.change(event.target.value),
			},
			control.options.map((option) =>
				createElement('option', { key: option.value, value: option.value }, option.label),
			),
		);
	return createElement(
		'button',
		{
			key: control.id,
			type: 'button',
			'aria-label': control.label,
			disabled: control.disabled,
			...(control.kind === 'toggle' ? { 'aria-pressed': control.pressed } : {}),
			onMouseDown: keepSelection,
			onClick: control.run,
		},
		control.label,
	);
}

/** The groups as a row of `role="group"` clusters; nodes are keyed so focus survives updates. */
export function renderControlGroups(groups: readonly ControlGroup[]): ReactNode[] {
	return groups.map((group, index) =>
		createElement('span', { key: index, role: 'group' }, ...group.map(renderControl)),
	);
}
