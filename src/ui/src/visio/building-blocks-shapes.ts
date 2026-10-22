import { installOfficeUiThemeFor } from '../theme';
import type { ViewerController } from './controller';
import { canvasServices, watchCanvasServices } from './canvas-services';
import { localizeRibbon } from './ribbon-localize';
import { createShapesWindow } from './shapes-window';
import { viewerStyles } from './styles';
import { canvasAndRibbonStyles, visioThemeAliases } from './styles/index';
import { registerViewerControls } from './office-ui';
import { wireStencilPane } from './viewer-stencil';

/** The pane fills whatever box the host gives it, instead of docking beside the canvas. */
const surfaceStyles = `
.shapes-pane{position:static;flex:1 1 auto;width:100%;height:100%;border-right:0;box-shadow:none}
.shapes-pane .pane-collapse{display:none}
:host{display:block}`;

/**
 * The Shapes window (Search shapes, the drawing's own stencil and the built-in stencils), mounted
 * on its own. Dragging a master onto the mounted canvas, or pressing Enter on it, creates the shape
 * through the same flow as the full viewer, so it obeys the same edit permission.
 */
export function mountDiagramShapes(
	host: HTMLElement,
	controller: ViewerController,
	options: { locale?: string | undefined } = {},
): () => void {
	const doc = host.ownerDocument;
	registerViewerControls();
	installOfficeUiThemeFor(host);
	const root = host.shadowRoot ?? host.attachShadow({ mode: 'open' });
	const style = doc.createElement('style');
	style.textContent = `${visioThemeAliases}${viewerStyles}${canvasAndRibbonStyles}${surfaceStyles}`;
	const pane = createShapesWindow(doc);
	const live = doc.createElement('div');
	live.setAttribute('role', 'status');
	live.setAttribute('aria-live', 'polite');
	live.style.cssText =
		'position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap';
	root.append(style, pane, live);
	if (options.locale) localizeRibbon(pane, options.locale);
	const stencil = wireStencilPane(pane, controller, (message) => {
		live.textContent = message;
		canvasServices(controller)?.announce(message);
	});
	// Drops land on whichever canvas is mounted on this controller, now or later.
	let detach: (() => void) | undefined;
	const stopWatching = watchCanvasServices(controller, (services) => {
		detach?.();
		detach = services ? stencil.attach(services.viewport) : undefined;
	});
	return () => {
		stopWatching();
		detach?.();
		stencil.dispose();
		root.replaceChildren();
	};
}
