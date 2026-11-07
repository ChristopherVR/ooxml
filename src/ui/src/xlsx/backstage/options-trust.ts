/**
 * Excel's File > Options > Trust Center: PowerPoint's Security settings section with the one
 * setting Excel has, "Download fonts that are not installed or embedded". When the host set
 * `googleFonts` the toggle is read-only and shows the host's value.
 */
import type { OfficeOptionCategory, OptionsContext } from '../../controls';

/** The option key of the font download toggle. */
export const FONT_DOWNLOADS_OPTION = 'allowFontDownloads';

export function trustCategory(context: OptionsContext, locked: boolean): OfficeOptionCategory {
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
						info: t(
							'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Excel.',
						),
						...(locked ? { disabled: t('This setting is managed by your organization') } : {}),
					},
				],
			},
		],
	};
}
