import { MAX_LATEX_LENGTH } from 'ooxml-core/math';

import { createInsertDialogFrame, type InsertDialogFrame } from '../../dialog/insert-dialog-frame';
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
	let frame: InsertDialogFrame | undefined;
	const close = () => frame?.dialog.close();
	const accept = () => {
		const draft = fields.draft;
		if (!draft || signal.aborted) return;
		insert(convertLatexToOmml(draft.latex));
		close();
	};
	const sync = (canInsert: boolean) =>
		frame?.sync(
			t('pptx.draw.inkToMath'),
			t('pptx.common.close'),
			t('pptx.equation.cancel'),
			t('pptx.equation.insert'),
			canInsert,
		);
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
		onChange: (draft) => sync(Boolean(draft)),
	});
	frame = createInsertDialogFrame(root, fields.element, accept, () => {
		signal.removeEventListener('abort', close);
		frame?.dialog.remove();
	});
	frame.dialog.style.setProperty('--office-dialog-width', 'min(32rem, calc(100vw - 32px))');
	signal.addEventListener('abort', close, { once: true });
	sync(Boolean(fields.draft));
	frame.dialog.show();
	return true;
}
