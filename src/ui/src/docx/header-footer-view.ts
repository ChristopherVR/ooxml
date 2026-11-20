import type { DocumentModel, HeaderFooterContent, HeaderFooterSlots } from 'ooxml-core/docx';
import { translateUiText } from './localization';
import { effectiveSlots } from './header-footer-link';
import { renderStoryPreview } from './story-render';
import type { PictureSource } from './picture-paint';

function slot(
	root: HTMLElement,
	labelText: string,
	content: HeaderFooterContent | undefined,
	name: 'default' | 'first' | 'even',
	model: DocumentModel,
	pictureUrl: PictureSource | undefined,
): void {
	if (!content || !content.blocks.length) return;
	const wrap = document.createElement('div');
	wrap.className = 'dve-header-footer-slot';
	wrap.dataset.slot = name;
	const label = document.createElement('div');
	label.className = 'dve-header-footer-label';
	label.textContent = labelText;
	label.dataset.localeAriaLabel = labelText;
	const body = document.createElement('div');
	body.className = 'dve-header-footer-body';
	body.append(renderStoryPreview(content.blocks, model, 'all', pictureUrl));
	wrap.append(label, body);
	root.append(wrap);
}

function buildSlotsElement(
	slots: HeaderFooterSlots | undefined,
	kind: 'header' | 'footer',
	locale: string,
	model: DocumentModel,
	pictureUrl: PictureSource | undefined,
): HTMLElement | null {
	if (!slots || (!slots.default && !slots.first && !slots.even)) return null;
	const root = document.createElement('section');
	root.className = `dve-${kind}`;
	root.setAttribute('contenteditable', 'false');
	root.dataset.localeAriaLabel = kind === 'header' ? 'Header' : 'Footer';
	root.setAttribute('aria-label', translateUiText(root, kind === 'header' ? 'Header' : 'Footer'));
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'First page header' : 'First page footer'),
		slots.first,
		'first',
		model,
		pictureUrl,
	);
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'Even page header' : 'Even page footer'),
		slots.even,
		'even',
		model,
		pictureUrl,
	);
	slot(
		root,
		translateUiText(root, kind === 'header' ? 'Header' : 'Footer'),
		slots.default,
		'default',
		model,
		pictureUrl,
	);
	root.dataset.editorLocale = locale;
	return root;
}

/** The selected section drives the header/footer preview on this continuous surface. */
export function buildHeaderElement(
	model: DocumentModel,
	locale: string,
	sectionIndex = 0,
	pictureUrl?: PictureSource,
): HTMLElement | null {
	return buildSlotsElement(
		effectiveSlots(model, sectionIndex, 'headers'),
		'header',
		locale,
		model,
		pictureUrl,
	);
}
export function buildFooterElement(
	model: DocumentModel,
	locale: string,
	sectionIndex = 0,
	pictureUrl?: PictureSource,
): HTMLElement | null {
	return buildSlotsElement(
		effectiveSlots(model, sectionIndex, 'footers'),
		'footer',
		locale,
		model,
		pictureUrl,
	);
}
