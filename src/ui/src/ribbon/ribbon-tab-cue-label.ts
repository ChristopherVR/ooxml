import { normalizeLocale } from 'ooxml-core/i18n';

const cueLabels = {
	en: 'More tabs',
	fr: 'Autres onglets',
	de: 'Weitere Registerkarten',
	es: 'Más pestañas',
	'zh-CN': '更多选项卡',
	ja: 'その他のタブ',
	ko: '탭 더 보기',
};

const locales = Object.keys(cueLabels) as (keyof typeof cueLabels)[];

export const ribbonTabCueLabel = (locale: string) =>
	cueLabels[normalizeLocale(locale, locales, 'en')];
