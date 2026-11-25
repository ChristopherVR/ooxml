import type { EditorCore } from './editor-core';
import { imageNodeView } from './image-media';

export function bodyImageNodeViews(core: EditorCore) {
	return {
		image: imageNodeView(core.imageMedia, {
			editPicture: (pos) => core.inserts.pictureDialog.open(pos),
			editChart: () => core.formatDialogs.open('chartData'),
			maxWidth: () => core.contentWidth(),
			theme: () => core.model.theme,
			model: () => core.model,
			locale: () => core.locale,
		}),
	};
}
