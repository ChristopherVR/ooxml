import { createPhotoAlbum, PptxHandler } from 'ooxml-core/pptx';
import type { PhotoAlbumOptions, PhotoAlbumPicture } from 'ooxml-core/pptx';
import { downloadBlob } from '../export/download-helpers';
import { createImageElementFromFile } from '../render/image-file-insertion';
import { collectElementText } from '../render/slide-search';
import { createInsertDialogFrame } from '../../dialog/insert-dialog-frame';
import type { InsertText } from './ribbon-graphic-insert';
import type { OfficeUiCheckbox } from '../../form/checkbox';
import { definePptxCheckbox, definePptxSelect } from './office-aliases';

const MAX_SOURCE_BYTES = 64 * 1024 * 1024;

function selectOption(doc: Document, value: string, label: string): HTMLOptionElement {
	const option = doc.createElement('option');
	option.value = value;
	option.textContent = label;
	return option;
}

export function pickReuseSlides(host: HTMLElement, text: InsertText) {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	if (registry) definePptxCheckbox(registry);
	const body = doc.createElement('div');
	const file = doc.createElement('input');
	file.type = 'file';
	file.accept = '.pptx,.pptm,.ppsx';
	file.setAttribute('aria-label', text('pptx.insert.reuseSlides', 'Reuse Slides'));
	const status = doc.createElement('p');
	status.setAttribute('role', 'status');
	const list = doc.createElement('fieldset');
	list.style.cssText = 'max-height:45vh;overflow:auto';
	body.append(file, status, list);
	return new Promise<{ source: ArrayBuffer; indexes: number[] } | undefined>((resolve) => {
		let source: ArrayBuffer | undefined;
		let request = 0;
		const frame = createInsertDialogFrame(
			host,
			body,
			() => {
				const indexes = chosen();
				if (!source || !indexes.length) return;
				resolve({ source, indexes });
				frame.dialog.close();
			},
			() => {
				request += 1;
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		const chosen = () =>
			[...list.querySelectorAll<OfficeUiCheckbox>('pptx-ui-checkbox')]
				.filter((checkbox) => checkbox.checked)
				.map((checkbox) => Number(checkbox.value));
		const sync = () =>
			frame.sync(
				text('pptx.insert.reuseSlides', 'Reuse Slides'),
				text('pptx.common.close', 'Close'),
				text('pptx.common.cancel', 'Cancel'),
				text('pptx.common.insert', 'Insert'),
				!!source && chosen().length > 0,
			);
		list.addEventListener('change', sync);
		file.addEventListener('change', () => {
			const selected = file.files?.[0];
			const token = ++request;
			source = undefined;
			list.replaceChildren();
			status.textContent = '';
			sync();
			if (selected) void read(selected, token);
		});
		async function read(selected: File, token: number): Promise<void> {
			const handler = new PptxHandler();
			try {
				if (selected.size > MAX_SOURCE_BYTES) throw new Error('The source file is too large');
				const bytes = await selected.arrayBuffer();
				const data = await handler.load(bytes, {
					eagerDecodeImages: false,
					allowExternalImages: false,
				});
				if (token !== request) return;
				source = bytes;
				for (const [index, slide] of data.slides.entries()) {
					const checkbox = doc.createElement('pptx-ui-checkbox') as OfficeUiCheckbox;
					checkbox.value = String(index);
					const title =
						slide.name || slide.elements.map(collectElementText).find(Boolean)?.trim().slice(0, 96);
					const name = text('pptx.zoom.slideNumber', 'Slide {{number}}', { number: index + 1 });
					const label = doc.createElement('label');
					label.style.display = 'block';
					const slideLabel = title ? name + ' · ' + title : name;
					checkbox.setAttribute('aria-label', slideLabel);
					label.append(checkbox, ' ', slideLabel);
					list.append(label);
				}
				sync();
			} catch {
				if (token === request)
					status.textContent = text(
						'pptx.graphic.insertFailed',
						'Could not insert the selected item.',
					);
			} finally {
				handler.dispose();
			}
		}
		sync();
		frame.dialog.show();
	});
}

export function pickPhotoAlbum(host: HTMLElement, text: InsertText) {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	if (registry) {
		definePptxCheckbox(registry);
		definePptxSelect(registry);
	}
	const heading = text('pptx.insert.photoAlbum', 'Photo Album');
	const body = doc.createElement('div');
	const title = doc.createElement('input');
	title.value = heading;
	title.setAttribute('aria-label', text('pptx.insert.albumTitle', 'Album title'));
	const file = doc.createElement('input');
	file.type = 'file';
	file.accept = 'image/png,image/jpeg,image/gif,image/webp,image/bmp';
	file.multiple = true;
	file.setAttribute('aria-label', text('pptx.insert.pictures', 'Pictures'));
	const list = doc.createElement('pptx-ui-select');
	list.setAttribute('aria-label', text('pptx.insert.pictures', 'Pictures'));
	const row = doc.createElement('div');
	row.style.cssText = 'display:flex;flex-wrap:wrap;gap:8px';
	const layout = doc.createElement('pptx-ui-select');
	const layoutLabel = doc.createElement('label');
	const layoutName = text('pptx.insert.picturesPerSlide', 'Pictures per slide');
	layout.setAttribute('aria-label', layoutName);
	layoutLabel.append(layoutName, ' ', layout);
	layout.append(...['1', '2', '4'].map((value) => selectOption(doc, value, value)));
	layout.value = '1';
	const captions = doc.createElement('pptx-ui-checkbox') as OfficeUiCheckbox;
	const captionsLabel = doc.createElement('label');
	const captionsName = text('pptx.insert.albumCaptions', 'Captions below pictures');
	captions.setAttribute('aria-label', captionsName);
	captionsLabel.append(captions, ' ', captionsName);
	const up = doc.createElement('button');
	up.type = 'button';
	up.textContent = text('pptx.customShows.moveUp', 'Move Up');
	const down = doc.createElement('button');
	down.type = 'button';
	down.textContent = text('pptx.customShows.moveDown', 'Move Down');
	const remove = doc.createElement('button');
	remove.type = 'button';
	remove.textContent = text('pptx.common.remove', 'Remove');
	row.append(up, down, remove);
	const status = doc.createElement('p');
	status.setAttribute('role', 'status');
	body.style.cssText = 'display:grid;gap:8px';
	body.append(title, file, list, row, layoutLabel, captionsLabel, status);
	return new Promise<PhotoAlbumOptions | undefined>((resolve) => {
		const files: File[] = [];
		const abort = new AbortController();
		let working = false;
		const onClose = () => {
			abort.abort();
			resolve(undefined);
			frame.dialog.remove();
		};
		const frame = createInsertDialogFrame(
			host,
			body,
			() => {
				if (!working) void prepare();
			},
			onClose,
		);
		const sync = () => {
			const index = list.selectedIndex;
			up.disabled = working || index <= 0;
			down.disabled = working || index < 0 || index >= files.length - 1;
			remove.disabled = working || index < 0;
			for (const control of [file, list, title, layout, captions]) control.disabled = working;
			frame.sync(
				heading,
				text('pptx.common.close', 'Close'),
				text('pptx.common.cancel', 'Cancel'),
				text('pptx.insert.createAlbum', 'Create Album'),
				!working && !!title.value.trim() && files.length > 0,
			);
		};
		const repaint = (index: number) => {
			list.replaceChildren(
				...files.map((selected, position) => selectOption(doc, String(position), selected.name)),
			);
			list.selectedIndex = Math.min(index, files.length - 1);
			sync();
		};
		const move = (offset: number) => {
			const index = list.selectedIndex;
			const neighbor = index + offset;
			const selected = files[index];
			if (!selected || !files[neighbor]) return;
			files[index] = files[neighbor];
			files[neighbor] = selected;
			repaint(neighbor);
		};
		up.addEventListener('click', () => move(-1));
		down.addEventListener('click', () => move(1));
		remove.addEventListener('click', () => {
			if (list.selectedIndex < 0) return;
			files.splice(list.selectedIndex, 1);
			repaint(list.selectedIndex);
		});
		list.addEventListener('change', sync);
		title.addEventListener('input', sync);
		file.addEventListener('change', () => {
			const selected = Array.from(file.files ?? []);
			file.value = '';
			if (
				files.length + selected.length > 100 ||
				[...files, ...selected].reduce((sum, item) => sum + item.size, 0) > MAX_SOURCE_BYTES
			) {
				status.textContent = text(
					'pptx.graphic.insertFailed',
					'Could not insert the selected item.',
				);
				return;
			}
			status.textContent = '';
			files.push(...selected);
			repaint(files.length - selected.length);
		});
		async function prepare(): Promise<void> {
			working = true;
			sync();
			try {
				const pictures: PhotoAlbumPicture[] = [];
				for (const selected of files) {
					const image = await createImageElementFromFile(
						selected,
						{ width: 4096, height: 4096 },
						abort.signal,
					);
					if (abort.signal.aborted) return;
					if (!image?.imageData) throw new Error('Could not read picture');
					pictures.push({
						data: image.imageData,
						width: image.width,
						height: image.height,
						caption: selected.name.replace(/\.[^.]+$/u, ''),
					});
				}
				const perSlide = layout.value === '4' ? 4 : layout.value === '2' ? 2 : 1;
				resolve({
					title: title.value.trim(),
					pictures,
					picturesPerSlide: perSlide,
					captions: captions.checked,
				});
				frame.dialog.close();
			} catch {
				status.textContent = text(
					'pptx.graphic.insertFailed',
					'Could not insert the selected item.',
				);
			} finally {
				working = false;
				sync();
			}
		}
		sync();
		frame.dialog.show();
	});
}

export async function downloadPhotoAlbum(options: PhotoAlbumOptions): Promise<void> {
	const bytes = await createPhotoAlbum(options);
	const name = options.title.replace(/[<>:"/\\|?*]/gu, '_').slice(0, 100);
	downloadBlob(
		new Blob([bytes.slice()], {
			type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		}),
		name + '.pptx',
	);
}
