// The data-validation error alert (Stop / Warning / Information), a small modal inside the
// shadow root. Resolves with the button the user chose; Escape is Cancel.
import type { EditorContext } from '../context.js';
import { h } from './dom.js';

export type AlertStyle = 'stop' | 'warning' | 'information';
export type AlertAnswer = 'retry' | 'cancel' | 'yes' | 'no' | 'ok';

const BUTTONS: Record<AlertStyle, [AlertAnswer, string][]> = {
	stop: [
		['retry', 'Retry'],
		['cancel', 'Cancel'],
	],
	warning: [
		['yes', 'Yes'],
		['no', 'No'],
		['cancel', 'Cancel'],
	],
	information: [
		['ok', 'OK'],
		['cancel', 'Cancel'],
	],
};

export function showValidationAlert(
	ctx: EditorContext,
	style: AlertStyle,
	message: string,
	title?: string,
): Promise<AlertAnswer> {
	const doc = ctx.host.ownerDocument;
	const backdrop = h(doc, 'div', 'xg-alert-backdrop');
	backdrop.style.cssText =
		'position:fixed;inset:0;z-index:1000;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.08)';
	const box = h(doc, 'div', 'xg-popup xg-alert', { role: 'alertdialog', 'aria-modal': 'true' });
	box.style.position = 'relative';
	const heading = h(doc, 'h3');
	heading.textContent = title || ctx.t('Microsoft Excel');
	const text = h(doc, 'p');
	const prompt = style === 'warning' ? `\n\n${ctx.t('Continue?')}` : '';
	text.textContent = `${ctx.t(message)}${prompt}`;
	const actions = h(doc, 'div', 'xg-actions');
	box.append(heading, text, actions);
	backdrop.append(box);
	ctx.root.append(backdrop);
	box.setAttribute('aria-label', heading.textContent);
	return new Promise((resolve) => {
		const done = (answer: AlertAnswer) => {
			backdrop.remove();
			resolve(answer);
		};
		for (const [answer, label] of BUTTONS[style]) {
			const button = h(doc, 'button', '', { type: 'button', 'data-answer': answer });
			button.textContent = ctx.t(label);
			button.addEventListener('click', () => done(answer));
			actions.append(button);
		}
		box.addEventListener('keydown', (event) => {
			event.stopPropagation();
			if (event.key === 'Escape') done('cancel');
		});
		actions.querySelector<HTMLButtonElement>('button')?.focus();
	});
}
