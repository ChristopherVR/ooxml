import {
	WORD_ART_STYLES,
	wordArtTransforms,
	renderWordArtSvg,
	wordArtCssShadow,
	wordArtPaintEquals,
	type WordArtText,
} from 'ooxml-core/drawingml';
import { defineDialog, type OfficeUiDialog } from '../dialog/dialog';
import { defineDialogFooter, type OfficeDialogFooterState } from '../dialog/dialog-footer';
import css from './word-art-picker.css?raw';
import { wordArtTransformLabel } from './word-art-translations';

let previewId = 0;

export interface WordArtLabels {
	heading: string;
	text: string;
	placeholder: string;
	style: (index: number) => string;
	transform: string;
	preview: string;
	locale: string;
	confirm: string;
	cancel: string;
	close: string;
}

export function pickWordArt(
	host: HTMLElement | ShadowRoot,
	labels: WordArtLabels,
	initial?: WordArtText,
): Promise<WordArtText | undefined> {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	defineDialog(registry);
	defineDialogFooter(registry);
	const dialog = doc.createElement('office-ui-dialog') as OfficeUiDialog;
	dialog.heading = labels.heading;
	dialog.closeLabel = labels.close;
	const style = doc.createElement('style');
	style.textContent = css;
	const body = doc.createElement('div');
	body.className = 'office-word-art';
	const input = doc.createElement('input');
	input.type = 'text';
	input.setAttribute('aria-label', labels.text);
	input.value = initial?.text ?? labels.placeholder;
	input.addEventListener('focus', () => input.select(), { once: true });
	const choices = doc.createElement('div');
	choices.className = 'office-word-art-choices';
	const transformLabel = doc.createElement('label');
	transformLabel.textContent = labels.transform;
	const transform = doc.createElement('select');
	const initialPreset = initial?.style.warp?.preset ?? 'textNoShape';
	for (const preset of new Set([...wordArtTransforms(), initialPreset])) {
		const option = doc.createElement('option');
		option.value = preset;
		option.textContent = wordArtTransformLabel(preset, labels.locale);
		transform.append(option);
	}
	transform.value = initialPreset;
	transformLabel.append(transform);
	const preview = doc.createElement('div');
	preview.className = 'office-word-art-preview';
	preview.setAttribute('role', 'img');
	preview.setAttribute('aria-label', labels.preview);
	const id = `office-word-art-preview-${++previewId}`;
	const footer = doc.createElement('office-ui-dialog-footer') as HTMLElement & {
		state: OfficeDialogFooterState;
	};
	footer.slot = 'footer';
	body.append(input, choices, transformLabel, preview);
	dialog.append(style, body, footer);

	return new Promise((resolve) => {
		const variants = [...WORD_ART_STYLES];
		const initialStyle = initial?.style;
		let selected = initial
			? variants.findIndex((item) => wordArtPaintEquals(item, initial.style))
			: 0;
		if (initial && selected < 0) selected = variants.push(initial.style) - 1;
		const value = (): WordArtText | undefined => {
			const text = input.value.trim();
			const chosen = variants[selected];
			if (!chosen || !text) return undefined;
			const { warp: _previousWarp, ...paint } = chosen;
			const preset = transform.value;
			const warp =
				preset === 'textNoShape'
					? undefined
					: initialStyle?.warp?.preset === preset
						? initialStyle.warp
						: { preset, guides: [] };
			return { text, style: { ...paint, ...(warp ? { warp } : {}) } };
		};
		const submit = () => {
			const current = value();
			if (!current) return;
			resolve(current);
			dialog.close();
		};
		const sync = () => {
			const current = value();
			preview.innerHTML = current
				? renderWordArtSvg({
						id,
						style: current.style,
						lines: [current.text],
						width: 344,
						height: 80,
						fontPx: 32,
					})
				: '';
			for (const [index, choice] of Array.from(choices.querySelectorAll('button')).entries()) {
				choice.setAttribute('aria-pressed', String(index === selected));
				choice.tabIndex = index === selected ? 0 : -1;
			}
			footer.state = {
				actions: [
					{ id: 'cancel', label: labels.cancel },
					{
						id: 'confirm',
						label: labels.confirm,
						variant: 'primary',
						disabled: !input.value.trim(),
					},
				],
			};
		};
		variants.forEach((variant, index) => {
			const button = doc.createElement('button');
			button.type = 'button';
			button.setAttribute('aria-label', labels.style(index + 1));
			button.textContent = 'Aa';
			button.style.color = `#${variant.fill}`;
			if (variant.gradient) {
				button.style.backgroundImage = `linear-gradient(${variant.gradient.angle / 60000 + 90}deg, #${variant.fill}, #${variant.gradient.end})`;
				button.style.backgroundClip = 'text';
				button.style.webkitTextFillColor = 'transparent';
			}
			if (variant.outline) button.style.webkitTextStroke = `1px #${variant.outline}`;
			button.style.textShadow = wordArtCssShadow(variant, 0.7) ?? '';
			button.addEventListener('click', () => {
				selected = index;
				sync();
			});
			choices.append(button);
		});
		input.addEventListener('input', sync);
		transform.addEventListener('change', sync);
		choices.addEventListener('keydown', (event) => {
			const columns =
				doc.defaultView?.getComputedStyle(choices).gridTemplateColumns.split(' ').length ?? 1;
			const delta = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -columns, ArrowDown: columns };
			const offset = Object.entries(delta).find(([key]) => key === event.key)?.[1];
			if (offset === undefined && event.key !== 'Home' && event.key !== 'End') return;
			event.preventDefault();
			selected =
				event.key === 'Home'
					? 0
					: event.key === 'End'
						? variants.length - 1
						: Math.max(0, Math.min(variants.length - 1, selected + (offset ?? 0)));
			sync();
			choices.querySelectorAll('button')[selected]?.focus();
		});
		input.addEventListener('keydown', (event) => {
			if (event.key === 'Enter' && !event.isComposing) {
				event.preventDefault();
				submit();
			}
		});
		footer.addEventListener('office-dialog-footer-request', (event) => {
			if ((event as CustomEvent<{ id: string }>).detail.id === 'confirm') submit();
			else dialog.close();
		});
		dialog.addEventListener('office-dialog-close', () => {
			resolve(undefined);
			queueMicrotask(() => dialog.remove());
		});
		sync();
		host.append(dialog);
		dialog.show();
	});
}
