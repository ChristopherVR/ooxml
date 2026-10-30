import { parseVsdx } from 'ooxml-core/visio';
import { VISIO_TEMPLATES, createTemplateVsdx, type VisioTemplateId } from 'ooxml-core/visio/ui';
import { exportPageSvg } from './export-svg';

const cache = new Map<VisioTemplateId, Promise<string | undefined>>();

/**
 * The first page of a New template drawn by the viewer's own page renderer: the template is built
 * as a real drawing, parsed and exported as a static SVG, returned as an image data URL. Nothing
 * is hand drawn, so a tile always matches what the template creates. Undefined when it cannot be
 * rendered (no DOM, or a failure), in which case the tile keeps its text only.
 */
export function templateThumbnail(id: VisioTemplateId): Promise<string | undefined> {
	let found = cache.get(id);
	if (!found) {
		found = (async () => {
			try {
				const model = await parseVsdx(await createTemplateVsdx(id));
				const { svg } = exportPageSvg(model, 0);
				return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
			} catch {
				return undefined;
			}
		})();
		cache.set(id, found);
	}
	return found;
}

/** Fill the thumbnail of every New tile under `root`, once; later calls reuse the cached images. */
export async function fillTemplateThumbnails(root: ParentNode): Promise<void> {
	await Promise.all(
		VISIO_TEMPLATES.map(async (template) => {
			const action = template.id === 'blank' ? 'new-blank' : `new-template:${template.id}`;
			const tile = root.querySelector<HTMLElement>(`[data-backstage-action="${action}"]`);
			if (!tile || tile.querySelector('.backstage-thumb')) return;
			const source = await templateThumbnail(template.id);
			if (!source || tile.querySelector('.backstage-thumb')) return;
			const image = tile.ownerDocument.createElement('img');
			image.className = 'backstage-thumb';
			image.alt = '';
			image.setAttribute('aria-hidden', 'true');
			image.src = source;
			tile.prepend(image);
		}),
	);
}
