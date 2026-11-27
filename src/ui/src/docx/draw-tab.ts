// The Draw tab, loaded when it is first selected (or a Draw command first runs): its strings, its
// panel and its commands. `draw-loader.ts` is the editor's only static link to it.
import { strings } from './localization-strings';
import { deDraw, enDraw, esDraw, frDraw, zhDraw } from './locales/draw';
import { buildDrawPanel, defineDrawElements } from './draw-ribbon';
import type { EditorCore } from './editor-core';
import { localizeElement } from './localization';
import { refitRibbon } from './ribbon-overflow';
import { injectDrawStyles } from './draw-replay-run';
import { attachRibbonHistoryKeys } from './ribbon-history-keys';
import { emit } from './events';
import { registerDrawIcons } from './draw-icons';

for (const [table, entries] of [
	[strings.en, enDraw],
	[strings.fr, frDraw],
	[strings.de, deDraw],
	[strings.es, esDraw],
	[strings['zh-CN'], zhDraw],
] as const)
	Object.assign(table as Record<string, string>, entries);

/** Builds the Draw panel once, in the editor's language, and defines its pen strip. */
export function mountDrawPanel(core: EditorCore): void {
	registerDrawIcons();
	defineDrawElements();
	const toolbar = core.shell.toolbar;
	const panel = toolbar?.querySelector<HTMLElement>('[data-ribbon-tab="draw"]');
	if (!toolbar || !panel || panel.childElementCount) return;
	injectDrawStyles(toolbar);
	// After a Draw command (Ink to Shape) focus stays on the ribbon, where Ctrl+Z still undoes.
	attachRibbonHistoryKeys(toolbar, (key) =>
		emit(toolbar, 'ribbon-action', { type: 'history', key }),
	);
	buildDrawPanel(panel);
	localizeElement(panel, core.locale);
	refitRibbon(toolbar);
}

export { destroyDraw, handleDrawAction, resetDraw, syncDraw } from './draw-actions';
