import backstage from './backstage.css?raw';
import canvas from './canvas.css?raw';
import ribbon from './ribbon.css?raw';
import shapes from './shapes.css?raw';
import sizePosition from './size-position.css?raw';
import textTool from './text-tool.css?raw';
import paintProperties from './paint-properties.css?raw';
import presentation from './presentation.css?raw';
import { visioThemeAliases, visioThemeBridge } from './theme';

/**
 * Viewer stylesheets as text for the shadow root. New styles are CSS files here; the legacy
 * `styles.ts` string moves into this directory when it is next changed.
 */
export const canvasAndRibbonStyles = [
	ribbon,
	shapes,
	backstage,
	canvas,
	sizePosition,
	textTool,
	paintProperties,
	presentation,
	visioThemeBridge,
].join('\n');

export { visioThemeAliases };
