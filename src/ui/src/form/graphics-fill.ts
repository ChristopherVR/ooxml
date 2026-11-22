import type { DrawingPictureEffect, PictureEffectPatch } from 'ooxml-core/drawingml';

import { createInsertDialogFrame } from '../dialog/insert-dialog-frame';

interface GraphicsFillLabels {
	heading: string;
	color: string;
	reset: string;
	cancel: string;
	close: string;
	apply: string;
}

export function pickGraphicsFill(
	host: HTMLElement | ShadowRoot,
	effects: readonly DrawingPictureEffect[] | undefined,
	labels: GraphicsFillLabels,
): Promise<PictureEffectPatch | undefined> {
	const doc = host.ownerDocument;
	const body = doc.createElement('div');
	const label = doc.createElement('label');
	const color = doc.createElement('input');
	color.type = 'color';
	const fill = effects?.find((effect) => effect.kind === 'standard' && effect.name === 'clrRepl');
	const current = fill?.kind === 'standard' ? fill.colors?.[0] : undefined;
	if (current?.kind === 'srgb' && /^[0-9a-f]{6}$/i.test(current.value)) {
		color.value = '#' + current.value;
	}
	label.append(labels.color, ' ', color);
	const reset = doc.createElement('button');
	reset.type = 'button';
	reset.textContent = labels.reset;
	reset.disabled = !fill;
	reset.style.marginInlineStart = '12px';
	body.append(label, reset);
	return new Promise((resolve) => {
		const frame = createInsertDialogFrame(
			host,
			body,
			() => {
				resolve({
					standard: {
						name: 'clrRepl',
						colors: [{ kind: 'srgb', value: color.value.slice(1), transforms: [] }],
					},
				});
				frame.dialog.close();
			},
			() => {
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		reset.addEventListener('click', () => {
			resolve({ removeStandard: 'clrRepl' });
			frame.dialog.close();
		});
		frame.sync(labels.heading, labels.close, labels.cancel, labels.apply, true);
		frame.dialog.show();
	});
}
