import type { PptxElement, TextPptxElement } from 'ooxml-core/pptx';
import { defineDialog, type OfficeUiDialog } from '../../dialog/dialog';
import { defineSymbolPicker } from '../../form/symbol-picker';
import { pickOfficeIcon } from '../../form/icon-picker';
import { renderLocalSvg } from '../../form/local-svg';
import { generateElementId } from '../render/element-clipboard';
import { captureInlineSymbolCaret } from '../render/inline-symbol';
import { WORDART_STYLES } from '../render/ribbon-galleries/wordart-styles-gallery';
import { wordArtTileSvg } from '../render/ribbon-galleries/wordart-preview-svg';
import { wordArtBodyStyle, wordArtRunStyle } from '../render/ribbon-galleries/wordart-run-style';
import {
	type GraphicInsertIntent,
	type RibbonGraphicTarget,
	type RibbonInsertState,
	insertLabel,
} from '../render/ribbon-insert-state';
import { createInsertDialogFrame } from './insert-dialog-frame';
import { pickRibbonZoom } from './ribbon-zoom-picker';
import css from './ribbon-graphic-insert.css?raw';

export type InsertText = (
	key: string,
	fallback: string,
	params?: Record<string, string | number>,
) => string;

export function createGraphicInsert(
	host: HTMLElement,
	readState: () => RibbonInsertState,
): (intent: GraphicInsertIntent) => void {
	let busy = false;
	return (intent) => {
		if (busy) return;
		const state = readState();
		const target = state.getGraphicTarget?.();
		const slide = target?.slides[target.slideIndex];
		if (!state.editable || !target || !slide) return;
		busy = true;
		const inlineSymbol =
			intent.kind === 'graphic' && intent.value === 'symbol'
				? captureInlineSymbolCaret(host.ownerDocument)
				: undefined;
		const text: InsertText = (key, fallback, params) => insertLabel(state, [key, fallback], params);
		const current = () => {
			const liveState = readState();
			const live = liveState.getGraphicTarget?.();
			return host.isConnected && liveState.editable && live?.slides[live.slideIndex] === slide
				? live
				: undefined;
		};
		const picked =
			intent.kind === 'zoom'
				? pickRibbonZoom(host, target, intent.value, text, current)
				: inlineSymbol
					? pickSymbol(host, text('pptx.insert.symbol', 'Symbol'), text).then((symbol) => {
							inlineSymbol(current() ? symbol : undefined);
							return undefined;
						})
					: pickGraphic(host, target, intent.value, text);
		void picked
			.then((element) => {
				if (element) current()?.insert(element);
			})
			.catch(() => {
				inlineSymbol?.();
				if (!host.isConnected) return;
				const alert = host.ownerDocument.createElement('div');
				alert.setAttribute('role', 'alert');
				alert.className = 'pptx-graphic-error';
				alert.textContent = text(
					'pptx.graphic.insertFailed',
					'Could not insert the selected item.',
				);
				host.append(alert);
				setTimeout(() => alert.remove(), 5000);
			})
			.finally(() => {
				busy = false;
			});
	};
}

function newGraphicText(
	target: RibbonGraphicTarget,
	name: string,
	content: string,
	fontSize: number,
): TextPptxElement {
	const { width, height } = target.canvasSize;
	return {
		type: 'text',
		id: generateElementId(),
		name,
		text: content,
		x: width / 4,
		y: height / 3,
		width: width / 2,
		height: height / 3,
		textStyle: { fontSize, color: '#000000', align: 'center' },
	};
}

async function pickGraphic(
	host: HTMLElement,
	target: RibbonGraphicTarget,
	kind: 'icons' | 'symbol' | 'wordArt',
	text: InsertText,
): Promise<PptxElement | undefined> {
	if (kind === 'symbol') {
		const heading = text('pptx.insert.symbol', 'Symbol');
		const symbol = await pickSymbol(host, heading, text);
		return symbol ? newGraphicText(target, heading, symbol, 24) : undefined;
	}
	if (kind === 'wordArt') return pickWordArt(host, target, text);

	const picked = await pickOfficeIcon(host, {
		heading: text('pptx.insert.icons', 'Icons'),
		search: text('pptx.graphic.searchIcons', 'Search icons'),
		color: text('pptx.graphic.iconColor', 'Icon color'),
		empty: text('pptx.graphic.noIcons', 'No matching icons.'),
		cancel: text('pptx.common.cancel', 'Cancel'),
		insert: text('pptx.common.insert', 'Insert'),
		close: text('pptx.common.close', 'Close'),
	});
	if (!picked) return undefined;
	const fallback = await renderLocalSvg(host.ownerDocument, picked.svg, 192, 192);
	const size = Math.min(target.canvasSize.width, target.canvasSize.height) / 5;
	return {
		type: 'image',
		id: generateElementId(),
		name: picked.name,
		altText: picked.name,
		x: (target.canvasSize.width - size) / 2,
		y: (target.canvasSize.height - size) / 2,
		width: size,
		height: size,
		imageData: fallback.toDataURL('image/png'),
		svgData: `data:image/svg+xml;base64,${btoa(picked.svg)}`,
	};
}

function pickSymbol(
	host: HTMLElement,
	heading: string,
	text: InsertText,
): Promise<string | undefined> {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	defineDialog(registry);
	defineSymbolPicker(registry);
	const dialog = doc.createElement('office-ui-dialog') as OfficeUiDialog;
	dialog.heading = heading;
	dialog.closeLabel = text('pptx.common.close', 'Close');
	const picker = doc.createElement('office-ui-symbol-picker');
	picker.setAttribute('label', heading);
	dialog.append(picker);
	return new Promise((resolve) => {
		picker.addEventListener('office-symbol-pick', (event) => {
			resolve((event as CustomEvent<{ symbol: string }>).detail.symbol);
			dialog.close();
		});
		dialog.addEventListener('office-dialog-close', () => {
			resolve(undefined);
			queueMicrotask(() => dialog.remove());
		});
		host.append(dialog);
		dialog.show();
	});
}

function pickWordArt(host: HTMLElement, target: RibbonGraphicTarget, text: InsertText) {
	const doc = host.ownerDocument;
	const heading = text('pptx.insert.wordArt', 'WordArt');
	const body = doc.createElement('div');
	body.className = 'pptx-wordart-insert';
	const style = doc.createElement('style');
	style.textContent = css;
	const input = doc.createElement('input');
	input.type = 'text';
	input.setAttribute('aria-label', text('pptx.graphic.wordArtText', 'WordArt text'));
	input.value = text('pptx.graphic.wordArtPlaceholder', 'Your text here');
	input.addEventListener('focus', () => input.select(), { once: true });
	const grid = doc.createElement('div');
	grid.className = 'pptx-wordart-grid';
	body.append(style, input, grid);
	return new Promise<PptxElement | undefined>((resolve) => {
		let selected: (typeof WORDART_STYLES)[number] | undefined;
		const frame = createInsertDialogFrame(
			host,
			body,
			() => {
				if (!selected || !input.value.trim()) return;
				const element = newGraphicText(target, heading, input.value, 40);
				element.textStyle = {
					...element.textStyle,
					...wordArtRunStyle(selected, target.themeColors),
					...wordArtBodyStyle(selected),
				};
				resolve(element);
				frame.dialog.close();
			},
			() => {
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		const sync = () =>
			frame.sync(
				heading,
				text('pptx.common.close', 'Close'),
				text('pptx.common.cancel', 'Cancel'),
				text('pptx.common.insert', 'Insert'),
				!!selected && !!input.value.trim(),
			);
		for (const [index, preset] of WORDART_STYLES.entries()) {
			const button = doc.createElement('button');
			button.type = 'button';
			button.title = text(`pptx.gallery.wordArtStyles.style${index}`, preset.label);
			button.setAttribute('aria-label', button.title);
			button.setAttribute('aria-pressed', 'false');
			button.innerHTML = wordArtTileSvg(
				`insert-wordart-${index}`,
				wordArtRunStyle(preset, target.themeColors),
				{ width: 44, height: 44 },
			);
			button.addEventListener('click', () => {
				grid.querySelector('[aria-pressed="true"]')?.setAttribute('aria-pressed', 'false');
				button.setAttribute('aria-pressed', 'true');
				selected = preset;
				sync();
			});
			grid.append(button);
		}
		input.addEventListener('input', sync);
		sync();
		frame.dialog.show();
	});
}
