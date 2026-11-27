import { expect, it } from 'vitest';
import { ribbonTabCueLabel } from './ribbon-tab-cue-label';

it('uses the supported language for regional editor locales', () => {
	expect(ribbonTabCueLabel('fr-FR')).toBe('Autres onglets');
	expect(ribbonTabCueLabel('zh-Hans-CN')).toBe('更多选项卡');
	expect(ribbonTabCueLabel('en-US')).toBe('More tabs');
});
