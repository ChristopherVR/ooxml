import { MAX_LATEX_LENGTH } from 'ooxml-core/math';

import { createEquationFields } from '../../math/equation-fields';
import { convertLatexToOmml } from '../render/latex-to-omml';
import type { PptxDrawHostOptions } from '../render/draw-canvas';

export async function recognizeDrawEquation(
	options: Omit<Parameters<typeof openDrawEquation>[0], 'latex'> & {
		provider: NonNullable<PptxDrawHostOptions['recognizeMath']>;
		strokes: Parameters<NonNullable<PptxDrawHostOptions['recognizeMath']>>[0];
		isCurrent: () => boolean;
	},
): Promise<boolean> {
	const { provider, strokes, signal, isCurrent } = options;
	if (signal.aborted) {
		return false;
	}
	const latex = await provider(
		strokes.map((points) => points.map(({ x, y }) => ({ x, y }))),
		signal,
	);
	return Boolean(latex && isCurrent() && openDrawEquation({ ...options, latex }));
}

/** Mount the same equation fields Word and Excel use inside a modal dialog. */
export function openDrawEquation(options: {
	root: HTMLElement;
	latex: string;
	t: (key: string) => string;
	signal: AbortSignal;
	insert: (omml: Record<string, unknown>) => void;
}): boolean {
	const { root, latex, t, signal, insert } = options;
	if (
		typeof latex !== 'string' ||
		!latex.trim() ||
		latex.length > MAX_LATEX_LENGTH ||
		signal.aborted
	) {
		return false;
	}
	const doc = root.ownerDocument;
	const previousFocus = doc.activeElement;
	const dialog = doc.createElement('dialog');
	dialog.setAttribute('aria-label', t('pptx.draw.inkToMath'));
	Object.assign(dialog.style, {
		width: 'min(32rem, 90vw)',
		background: 'var(--office-surface, Canvas)',
		color: 'var(--office-text, CanvasText)',
		border: '1px solid var(--office-border, GrayText)',
	});
	const title = doc.createElement('h2');
	title.textContent = t('pptx.draw.inkToMath');
	const confirm = doc.createElement('button');
	confirm.type = 'button';
	confirm.textContent = t('pptx.equation.insert');
	const cancel = doc.createElement('button');
	cancel.type = 'button';
	cancel.textContent = t('pptx.equation.cancel');
	const accept = () => {
		const draft = fields.draft;
		if (!draft || signal.aborted) {
			return;
		}
		insert(convertLatexToOmml(draft.latex));
		dialog.close();
	};
	const fields = createEquationFields(doc, {
		labels: {
			equation: t('pptx.equation.insertTitle'),
			source: t('pptx.equation.latexInput'),
			display: t('pptx.equation.insertTitle'),
			preview: t('pptx.draw.preview'),
			empty: t('pptx.equation.previewPlaceholder'),
			invalid: t('pptx.draw.invalidEquation'),
		},
		latex,
		display: true,
		accept,
		onChange: (draft) => {
			confirm.disabled = !draft;
		},
	});
	confirm.addEventListener('click', accept);
	cancel.addEventListener('click', () => dialog.close());
	const close = () => dialog.close();
	signal.addEventListener('abort', close, { once: true });
	dialog.addEventListener(
		'close',
		() => {
			signal.removeEventListener('abort', close);
			dialog.remove();
			if (previousFocus instanceof HTMLElement && previousFocus.isConnected) {
				previousFocus.focus();
			}
		},
		{ once: true },
	);
	dialog.addEventListener('keydown', (event) => event.stopPropagation());
	dialog.append(title, fields.element, cancel, confirm);
	root.append(dialog);
	dialog.showModal();
	fields.source.focus();
	return true;
}
