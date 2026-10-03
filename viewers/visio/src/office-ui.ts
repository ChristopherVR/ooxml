import { registerControls } from 'ooxml-ui/controls';
import { getIcon, registerIcon } from 'ooxml-ui/icons';

/**
 * Visio-only glyphs (20x20 path data). Neutral Office glyphs come from ooxml-ui; add a glyph
 * there instead when another Office product could use it.
 */
const VISIO_ICONS: Readonly<Record<string, string>> = {
	visioLayers: 'm2.5 7 7.5-4 7.5 4-7.5 4ZM2.5 10.5l7.5 4 7.5-4M2.5 14l7.5 4 7.5-4',
	visioShapeData: 'M3 3.5h14v13H3ZM3 7.5h14M8 7.5v9M10.5 11h4M10.5 14h4',
	visioPagesPane: 'M2.5 3h15v14h-15ZM7 3v14M3.8 6h1.9M3.8 9h1.9M3.8 12h1.9',
	visioInspectorPane: 'M2.5 3h15v14h-15ZM13 3v14M14.5 6h1.5M14.5 9h1.5M14.5 12h1.5',
};

/** Define the shared Office controls the viewer chrome uses. Idempotent; browser only. */
export function registerViewerControls(registry?: CustomElementRegistry): void {
	for (const [name, path] of Object.entries(VISIO_ICONS))
		if (!getIcon(name)) registerIcon(name, path);
	registerControls(registry);
}
