import { createInsertDialogFrame } from '../../dialog/insert-dialog-frame';
import { type RibbonViewState, type RibbonViewZoom, viewLabel } from '../render/ribbon-view-state';
import { ZOOM_MAX_PERCENT, ZOOM_MIN_PERCENT } from '../render/zoom-step';

export function openViewZoom(
	host: HTMLElement,
	state: RibbonViewState,
	zoom: RibbonViewZoom,
): void {
	if (host.querySelector('office-ui-dialog[data-view-zoom]')) return;
	const doc = host.ownerDocument;
	const text = (key: string, fallback: string) => viewLabel(state, key, fallback);
	const body = doc.createElement('div');
	const label = doc.createElement('label');
	const percent = doc.createElement('input');
	percent.type = 'number';
	percent.required = true;
	percent.min = String(ZOOM_MIN_PERCENT);
	percent.max = String(ZOOM_MAX_PERCENT);
	percent.value = String(Math.round(zoom.getPercent()));
	percent.style.width = '8ch';
	label.append(text('pptx.view.zoomPercent', 'Zoom percentage'), ' ', percent);
	const presets = doc.createElement('div');
	presets.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px;margin-block:12px';
	let fit = false;
	const frame = createInsertDialogFrame(host, body, apply, () => frame.dialog.remove());
	frame.dialog.dataset.viewZoom = '';
	frame.dialog.style.setProperty('--office-dialog-width', 'min(380px, calc(100vw - 32px))');

	function apply(): void {
		if (fit) zoom.fit();
		else if (percent.validity.valid) zoom.setPercent(percent.valueAsNumber);
		else return;
		frame.dialog.close();
	}
	function sync(): void {
		for (const button of presets.querySelectorAll('button')) {
			button.setAttribute(
				'aria-pressed',
				String(fit ? button.dataset.percent === 'fit' : button.dataset.percent === percent.value),
			);
		}
		frame.sync(
			text('pptx.slideSorter.zoom', 'Zoom'),
			text('pptx.common.close', 'Close'),
			text('pptx.common.cancel', 'Cancel'),
			text('pptx.common.ok', 'OK'),
			fit || percent.validity.valid,
		);
	}
	for (const preset of [400, 200, 100, 75, 50, 33, 25, 'fit'] as const) {
		const button = doc.createElement('button');
		button.type = 'button';
		button.dataset.percent = String(preset);
		button.textContent =
			preset === 'fit' ? text('pptx.view.zoomToFit', 'Zoom to Fit') : preset + '%';
		button.addEventListener('click', () => {
			fit = preset === 'fit';
			percent.value = fit ? '100' : String(preset);
			sync();
		});
		presets.append(button);
	}
	percent.addEventListener('input', () => {
		fit = false;
		sync();
	});
	percent.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			apply();
		}
	});
	body.append(label, presets);
	sync();
	frame.dialog.show();
}
