import type { OfficeOptionCategory, OfficeOptionControl } from '../controls';
import { translateRibbonText, type VisioLocale } from './ribbon-locale';

/** The Trust Center toggle's key in the options values and the saved behaviour. */
export const FONT_DOWNLOADS_KEY = 'allowFontDownloads';

const FONT_DOWNLOADS_INFO =
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Visio.';

/**
 * Visio's File > Options > Trust Center, PowerPoint's wording with one setting: whether fonts
 * that are neither installed nor embedded are downloaded. `locked` is the host's decision
 * (`fontDownloadLock(googleFonts)`): when set, the toggle is read-only and says why.
 */
export function visioTrustCategory(
	locale: VisioLocale = 'en',
	locked?: boolean,
): OfficeOptionCategory {
	const t = (text: string) => translateRibbonText(locale, text);
	const toggle: OfficeOptionControl = {
		kind: 'toggle',
		key: FONT_DOWNLOADS_KEY,
		label: t('Download fonts that are not installed or embedded (Google Fonts)'),
		info: t(FONT_DOWNLOADS_INFO),
		...(locked === undefined
			? {}
			: { disabled: t('This setting is managed by your organization') }),
	};
	return {
		id: 'trust',
		label: t('Trust Center'),
		description: t('Help keep your documents safe and secure.'),
		sections: [
			{
				id: 'trustSettings',
				title: t('Security settings'),
				description: t(
					'These settings help keep the content you open safe. We recommend the defaults.',
				),
				controls: [toggle],
			},
		],
	};
}
