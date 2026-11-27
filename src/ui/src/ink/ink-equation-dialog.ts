// The Ink Equation dialog Word's Ink Equation and Ink to Math and Excel's Ink to Math open, like
// Office's Math Input Control: a pad to write on, the host's recognizer turning what is written
// into LaTeX a moment after each stroke, and the shared equation fields to check and correct it
// before Insert. Loaded on first use. Only the strokes go to the recognizer.
import { translateInkText, type InkText } from 'ooxml-core/i18n';
import { createInsertDialogFrame, type InsertDialogFrame } from '../dialog/insert-dialog-frame';
import { createEquationFields, type EquationDraft } from '../math/equation-fields';
import { defineInkCanvas, type InkStrokeDetail, type OfficeUiInkCanvas } from './ink-canvas';
import type { InkMathRecognizer } from './math-recognizer';

export interface InkEquationDialogOptions {
	/** Where the dialog goes (the editor's element). */
	host: HTMLElement | ShadowRoot;
	recognizer: InkMathRecognizer;
	/** The product's locale: the dialog translates itself. */
	locale: string | undefined;
	heading: Extract<InkText, 'Ink Equation' | 'Ink to Math'>;
	/** Inserts the equation; false keeps the dialog open. */
	insert(draft: EquationDraft): boolean | void;
	/** After the dialog closed, inserted or not. */
	closed?(): void;
}

/** How long after the last stroke the writing is recognized. */
const SETTLE_MS = 450;
const PAD_HEIGHT = 180;
const SVG_NS = 'http://www.w3.org/2000/svg';

export function openInkEquationDialog(options: InkEquationDialogOptions): { close(): void } {
	const { host, recognizer, locale } = options;
	const doc = host.ownerDocument;
	const t = (text: InkText) => translateInkText(text, locale);
	const strokes: { x: number; y: number }[][] = [];
	let request: AbortController | undefined;
	let timer: ReturnType<typeof setTimeout> | undefined;
	// Declared first: the fields report their first draft while they are built.
	let frame: InsertDialogFrame | undefined;

	const body = doc.createElement('div');
	body.className = 'office-ink-equation';
	const pad = doc.createElement('div');
	pad.className = 'office-ink-equation-pad';
	Object.assign(pad.style, {
		position: 'relative',
		height: `${PAD_HEIGHT}px`,
		marginBottom: '8px',
		border: '1px solid var(--office-border, GrayText)',
		borderRadius: '4px',
		// A sheet of paper, as Office's pad is, so the black ink reads in every theme.
		background: '#ffffff',
		overflow: 'hidden',
	});
	const written = doc.createElementNS(SVG_NS, 'svg');
	written.setAttribute('aria-hidden', 'true');
	Object.assign(written.style, { position: 'absolute', inset: '0', width: '100%', height: '100%' });
	defineInkCanvas(doc.defaultView?.customElements);
	const canvas = doc.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
	canvas.tool = 'pen';
	canvas.color = '#000000';
	canvas.width = 2;
	canvas.touch = true;
	canvas.label = t('Write math here');
	pad.append(written, canvas);
	const toolbar = doc.createElement('div');
	Object.assign(toolbar.style, { display: 'flex', alignItems: 'center', gap: '8px' });
	const clear = doc.createElement('button');
	clear.type = 'button';
	clear.className = 'office-ink-equation-clear';
	clear.textContent = t('Clear');
	const status = doc.createElement('span');
	status.setAttribute('role', 'status');
	toolbar.append(clear, status);
	const fields = createEquationFields(doc, {
		labels: {
			equation: t('Equation'),
			source: t('LaTeX equation'),
			display: '',
			preview: t('Preview'),
			empty: t('Write an equation to see a preview.'),
			invalid: t('This LaTeX could not be converted to an equation.'),
		},
		display: true,
		accept,
		onChange: (draft) => sync(Boolean(draft)),
	});
	body.append(pad, toolbar, fields.element);

	function setLatex(latex: string): void {
		fields.source.value = latex;
		fields.source.dispatchEvent(new Event('input'));
	}
	function stop(): void {
		if (timer !== undefined) clearTimeout(timer);
		timer = undefined;
		request?.abort();
		request = undefined;
	}
	async function recognize(): Promise<void> {
		stop();
		if (!strokes.length) return;
		const abort = new AbortController();
		request = abort;
		status.textContent = t('Recognizing...');
		let latex: string | undefined;
		try {
			latex = await recognizer(
				strokes.map((points) => points.map(({ x, y }) => ({ x, y }))),
				abort.signal,
			);
		} catch {
			latex = undefined;
		}
		if (abort.signal.aborted || request !== abort) return;
		request = undefined;
		const text = typeof latex === 'string' ? latex.trim() : '';
		status.textContent = text
			? ''
			: t('The handwriting could not be recognized. Write it again or correct the LaTeX.');
		if (text) setLatex(text);
	}
	canvas.addEventListener('office-ink-stroke', (event) => {
		const points = (event as CustomEvent<InkStrokeDetail>).detail.samples.map(({ x, y }) => ({
			x,
			y,
		}));
		if (points.length === 0) return;
		strokes.push(points);
		const line = doc.createElementNS(SVG_NS, 'polyline');
		line.setAttribute('points', points.map(({ x, y }) => `${x},${y}`).join(' '));
		for (const [name, value] of [
			['fill', 'none'],
			['stroke', '#000000'],
			['stroke-width', '2'],
			['stroke-linecap', 'round'],
			['stroke-linejoin', 'round'],
		] as const)
			line.setAttribute(name, value);
		written.append(line);
		stop();
		timer = setTimeout(() => void recognize(), SETTLE_MS);
	});
	clear.addEventListener('click', () => {
		stop();
		strokes.length = 0;
		written.replaceChildren();
		status.textContent = '';
		setLatex('');
	});

	function accept(): void {
		const draft = fields.draft;
		if (!draft || options.insert(draft) === false) return;
		frame?.dialog.close();
	}
	function sync(canInsert: boolean): void {
		frame?.sync(t(options.heading), t('Close'), t('Cancel'), t('Insert'), canInsert);
	}
	frame = createInsertDialogFrame(host, body, accept, () => {
		stop();
		frame?.dialog.remove();
		options.closed?.();
	});
	frame.dialog.style.setProperty('--office-dialog-width', 'min(36rem, calc(100vw - 32px))');
	sync(Boolean(fields.draft));
	frame.dialog.show();
	return { close: () => frame?.dialog.close() };
}
