import canvas from './canvas.css?inline';
import ribbon from './ribbon.css?inline';

/**
 * Viewer stylesheets as text for the shadow root. New styles are CSS files here; the legacy
 * `styles.ts` string moves into this directory when it is next changed.
 */
export const canvasAndRibbonStyles = [ribbon, canvas].join('\n');
