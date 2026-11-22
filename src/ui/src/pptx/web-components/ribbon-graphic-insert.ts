import type { PptxElement, TextPptxElement } from 'ooxml-core/pptx';
import { pickOfficeSymbol } from '../../form/symbol-picker';
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
	type RibbonInsertWorkflow,
	type RibbonInsertState,
	insertLabel,
} from '../render/ribbon-insert-state';
import { createInsertDialogFrame } from '../../dialog/insert-dialog-frame';
import { pickScreenshot, pickScreenRecording } from './ribbon-capture';
import { downloadPhotoAlbum, pickPhotoAlbum, pickReuseSlides } from './ribbon-deck-insert';
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
	onInserted?: () => void,
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
			intent.kind === 'workflow'
				? runInsertWorkflow(host, target, intent.value, text, current, onInserted)
				: intent.kind === 'zoom'
					? pickRibbonZoom(host, target, intent.value, text, current, onInserted)
					: inlineSymbol
						? pickSymbol(host, text('pptx.insert.symbol', 'Symbol'), text).then((symbol) => {
								const content = current() ? symbol : undefined;
								if (inlineSymbol(content)) onInserted?.();
								return undefined;
							})
						: intent.value === 'icons'
							? pickIcons(host, target, text)
							: pickGraphic(host, target, intent.value, text);
		void picked
			.then((picked) => {
				if (!picked) return;
				const live = current();
				const elements = Array.isArray(picked) ? picked : [picked];
				if (!live || !elements.length) return;
				for (const element of elements) live.insert(element);
				onInserted?.();
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

async function runInsertWorkflow(
	host: HTMLElement,
	target: RibbonGraphicTarget,
	workflow: RibbonInsertWorkflow,
	text: InsertText,
	current: () => RibbonGraphicTarget | undefined,
	onInserted?: () => void,
): Promise<PptxElement | undefined> {
	switch (workflow) {
		case 'screenshot':
		case 'screenClipping':
			return pickScreenshot(host, target.canvasSize, workflow === 'screenClipping', text);
		case 'screenRecording':
			return pickScreenRecording(host, target.canvasSize, text);
		case 'photoAlbum': {
			const options = await pickPhotoAlbum(host, text);
			if (!options || !current()) return undefined;
			await downloadPhotoAlbum(options);
			onInserted?.();
			return undefined;
		}
		case 'reuseSlides': {
			const picked = await pickReuseSlides(host, text);
			const live = current();
			if (!picked || !live?.importSlides) return undefined;
			await live.importSlides(picked.source, picked.indexes);
			onInserted?.();
			return undefined;
		}
	}
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
	kind: 'symbol' | 'wordArt',
	text: InsertText,
): Promise<PptxElement | undefined> {
	if (kind === 'symbol') {
		const heading = text('pptx.insert.symbol', 'Symbol');
		const symbol = await pickSymbol(host, heading, text);
		return symbol ? newGraphicText(target, heading, symbol, 24) : undefined;
	}
	return pickWordArt(host, target, text);
}

async function pickIcons(
	host: HTMLElement,
	target: RibbonGraphicTarget,
	text: InsertText,
): Promise<PptxElement[]> {
	const picked = await pickOfficeIcon(host, {
		heading: text('pptx.insert.icons', 'Icons'),
		search: text('pptx.graphic.searchIcons', 'Search icons'),
		color: text('pptx.graphic.iconColor', 'Icon color'),
		empty: text('pptx.graphic.noIcons', 'No matching icons.'),
		cancel: text('pptx.common.cancel', 'Cancel'),
		insert: text('pptx.common.insert', 'Insert'),
		close: text('pptx.common.close', 'Close'),
	});
	const { width, height } = target.canvasSize;
	const columns = Math.max(1, Math.ceil(Math.sqrt(picked.length)));
	const rows = Math.max(1, Math.ceil(picked.length / columns));
	const pitch = Math.min(width / (columns + 1), height / (rows + 1), height / 4);
	const size = pitch * 0.8;
	return Promise.all(
		picked.map(async (icon, index): Promise<PptxElement> => {
			const fallback = await renderLocalSvg(host.ownerDocument, icon.svg, 192, 192);
			return {
				type: 'image',
				id: generateElementId(),
				name: icon.name,
				altText: icon.name,
				x: (width - columns * pitch) / 2 + ((index % columns) + 0.1) * pitch,
				y: (height - rows * pitch) / 2 + (Math.floor(index / columns) + 0.1) * pitch,
				width: size,
				height: size,
				imageData: fallback.toDataURL('image/png'),
				svgData: `data:image/svg+xml;base64,${btoa(icon.svg)}`,
			};
		}),
	);
}

function pickSymbol(
	host: HTMLElement,
	heading: string,
	text: InsertText,
): Promise<string | undefined> {
	return pickOfficeSymbol(host, {
		heading,
		close: text('pptx.common.close', 'Close'),
		font: text('pptx.ribbon.font', 'Font'),
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
