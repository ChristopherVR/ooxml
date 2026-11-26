import { wireRibbonTabCue as wireSharedRibbonTabCue } from '../ribbon/ribbon-tab-cue';
import { ribbonLocale } from './ribbon-localize';
import { translateRibbonText } from './ribbon-locale';

export function wireRibbonTabCue(ribbon: HTMLElement): () => void {
	return wireSharedRibbonTabCue(ribbon, () =>
		translateRibbonText(ribbonLocale(ribbon), 'More tabs'),
	);
}
