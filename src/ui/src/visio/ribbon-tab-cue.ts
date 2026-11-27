import { wireRibbonTabCue as wireSharedRibbonTabCue } from '../ribbon/ribbon-tab-cue';
import { ribbonTabCueLabel } from '../ribbon/ribbon-tab-cue-label';
import { ribbonLocale } from './ribbon-localize';

export function wireRibbonTabCue(ribbon: HTMLElement): () => void {
	return wireSharedRibbonTabCue(ribbon, () => ribbonTabCueLabel(ribbonLocale(ribbon)));
}
