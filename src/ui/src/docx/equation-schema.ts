import type { NodeSpec } from 'prosemirror-model';
/** Read-only native MathML display. The source OMML stays in the atom for saving. */
export function equationDom(omml: string, display: boolean): HTMLElement {
	const span = document.createElement('span');
	span.dataset.docxEquation = '1';
	span.className = display ? 'dve-equation dve-equation-display' : 'dve-equation';
	span.contentEditable = 'false';
	span.title = 'Equation preview (display only; Word layout may differ)';
	const unavailable = () => {
		span.classList.add('dve-equation-unavailable');
		span.setAttribute('role', 'math');
		span.setAttribute('aria-label', 'Equation preview unavailable; source is preserved');
		span.textContent = 'Equation preview unavailable';
	};
	void import('../math/equation-render').then(({ equationPreview, renderEquation }) => {
		const preview = equationPreview(omml);
		const math = preview && renderEquation(span.ownerDocument, preview, display, 'Equation');
		if (math) span.append(math);
		else unavailable();
	}, unavailable);
	return span;
}

export const equationNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: true,
	attrs: { omml: { default: '' }, display: { default: false }, format: { default: null } },
	toDOM: (node) => equationDom(String(node.attrs.omml), Boolean(node.attrs.display)),
};
