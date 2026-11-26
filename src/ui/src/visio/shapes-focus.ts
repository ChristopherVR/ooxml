import { modalActiveElement } from '../dialog/focus';

function findByAttribute(
	root: HTMLElement,
	attribute: string,
	value: string,
): HTMLElement | undefined {
	return [...root.querySelectorAll<HTMLElement>(`[${attribute}]`)].find(
		(element) => element.getAttribute(attribute) === value,
	);
}

export function rememberMasterFocus(
	quick: HTMLElement,
	quickToggle: HTMLElement,
): (() => void) | undefined {
	const active = modalActiveElement(quick.ownerDocument);
	const focusedMaster = quick.contains(active) ? active?.getAttribute('data-master') : undefined;
	if (!focusedMaster) return;
	return () => {
		(findByAttribute(quick, 'data-master', focusedMaster) ?? quickToggle).focus({
			preventScroll: true,
		});
	};
}

export function rememberSectionFocus(
	sections: HTMLElement,
	more: HTMLElement,
): (() => void) | undefined {
	const active = modalActiveElement(sections.ownerDocument);
	const focusedSection = sections.contains(active)
		? active?.closest<HTMLElement>('[data-stencil]')?.dataset.stencil
		: undefined;
	if (!focusedSection) return;
	const master = active?.getAttribute('data-master');
	return () => {
		const section = findByAttribute(sections, 'data-stencil', focusedSection);
		let next: HTMLElement | null | undefined;
		if (section) {
			if (master)
				next =
					findByAttribute(section, 'data-master', master) ??
					section.querySelector<HTMLElement>('[data-stencil-actions]');
			else if (active?.hasAttribute('data-stencil-actions'))
				next = section.querySelector<HTMLElement>('[data-stencil-actions]');
			else if (active?.hasAttribute('data-close-stencil'))
				next = section.querySelector<HTMLElement>('[data-close-stencil]');
			next ??= section.querySelector<HTMLElement>('.stencil-title');
		}
		(next ?? more).focus({ preventScroll: true });
	};
}
