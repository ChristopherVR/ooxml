import { dialogButton } from './dialog-fields';
import { localizeElement, translate, type EditorLocale } from './localization';
import { createStyleTile, markSelectedTiles } from './style-gallery';
import type { StylesData } from './styles-popover';

export interface StylesPane {
	element: HTMLElement;
	readonly isOpen: boolean;
	open(): void;
	close(): void;
	/** Redraws the lists for the current document and selection (a no-op while closed). */
	render(data: StylesData, locale: EditorLocale): void;
}

/**
 * Word's Styles pane: a docked list of every paragraph and character style. Each row previews its
 * style; choosing one applies it to the selection without closing the pane.
 */
export function createStylesPane(onClose: () => void): StylesPane {
	const element = document.createElement('aside');
	element.className = 'dve-styles-pane';
	element.setAttribute('role', 'complementary');
	element.hidden = true;
	let latest: StylesData | undefined;
	let key = '';
	let locale: EditorLocale = 'en';
	const create = dialogButton('Create a Style');
	const modify = dialogButton('Modify Style');
	create.addEventListener('click', () => latest?.author(false));
	modify.addEventListener('click', () => latest?.author(true));
	const actions = document.createElement('div');
	actions.className = 'dve-styles-pane-actions';
	actions.append(create, modify);
	const close = () => {
		element.hidden = true;
		onClose();
	};
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') close();
	});
	return {
		element,
		get isOpen() {
			return !element.hidden;
		},
		open() {
			element.hidden = false;
			element.querySelector<HTMLElement>('[aria-pressed="true"], button')?.focus();
		},
		close,
		render(data, next) {
			if (element.hidden) return;
			locale = next;
			const catalogsChanged =
				latest?.model.paragraphStyles !== data.model.paragraphStyles ||
				latest?.model.characterStyles !== data.model.characterStyles;
			latest = data;
			const signature = JSON.stringify([
				data.allParagraph,
				data.allCharacter,
				next,
				data.selectedParagraph,
				data.selectedCharacter,
				data.disabled,
				data.canAuthor,
				data.canModify,
			]);
			const listsChanged = signature !== key || catalogsChanged;
			key = signature;
			if (listsChanged) {
				const header = document.createElement('div');
				header.className = 'dve-styles-pane-header';
				const title = document.createElement('h2');
				title.textContent = translate(locale, 'Styles');
				const closeButton = document.createElement('button');
				closeButton.type = 'button';
				closeButton.textContent = '×';
				closeButton.setAttribute('aria-label', translate(locale, 'Close'));
				closeButton.addEventListener('click', close);
				header.append(title, closeButton);
				const list = (
					heading: string,
					styles: typeof data.paragraph,
					kind: 'paragraph' | 'character',
				) => {
					const wrap = document.createElement('div');
					const label = document.createElement('h3');
					label.textContent = heading;
					const items = document.createElement('div');
					items.className = 'dve-styles-pane-list';
					items.append(
						...styles.map((style) =>
							createStyleTile(
								style,
								data.model,
								kind,
								kind === 'paragraph' ? data.chooseParagraph : data.chooseCharacter,
							),
						),
					);
					markSelectedTiles(
						items,
						kind === 'paragraph' ? data.selectedParagraph : data.selectedCharacter,
						data.disabled,
					);
					wrap.append(label, items);
					return wrap;
				};
				create.disabled = !data.canAuthor;
				modify.disabled = !data.canAuthor || !data.canModify;
				element.replaceChildren(
					header,
					actions,
					list(translate(locale, 'Paragraph styles'), data.allParagraph, 'paragraph'),
					...(data.allCharacter.length
						? [
								list(
									translate(locale, 'Character styles'),
									[{ id: '', name: translate(locale, 'No character style') }, ...data.allCharacter],
									'character',
								),
							]
						: []),
				);
				localizeElement(element, locale);
			}
		},
	};
}
