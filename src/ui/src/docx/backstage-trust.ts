import type { OfficeOptionCategory, OptionsContext } from '../controls';

/** The option key of the Trust Center's font download toggle. */
export const FONT_DOWNLOADS_OPTION = 'allowFontDownloads';

/** The toggle's info text (PowerPoint's, naming Word). */
export const FONT_DOWNLOADS_INFO =
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.';

/**
 * File > Options > Trust Center, worded as PowerPoint's: one toggle for downloading the fonts a
 * document uses that are neither installed nor embedded. `locked` is the host's `googleFonts`
 * decision (`fontDownloadLock`): the toggle then shows that value read-only, with the reason.
 */
export function trustCategory(
	context: OptionsContext,
	locked: boolean | undefined,
): OfficeOptionCategory {
	const { t } = context;
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
				controls: [
					{
						kind: 'toggle',
						key: FONT_DOWNLOADS_OPTION,
						label: t('Download fonts that are not installed or embedded (Google Fonts)'),
						info: t(FONT_DOWNLOADS_INFO),
						...(locked === undefined
							? {}
							: { disabled: t('This setting is managed by your organization') }),
					},
				],
			},
		],
	};
}
