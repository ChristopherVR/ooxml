import backstage from './backstage.css?inline';
import canvas from './canvas.css?inline';
import ribbon from './ribbon.css?inline';
import shapes from './shapes.css?inline';

/**
 * Viewer stylesheets as text for the shadow root. New styles are CSS files here; the legacy
 * `styles.ts` string moves into this directory when it is next changed.
 */
export const canvasAndRibbonStyles = [ribbon, shapes, backstage, canvas].join('\n');
