import { normalizeLocale } from 'ooxml-core/i18n';

const messages = {
	en: 'Could not load add-ins.',
	fr: 'Impossible de charger les compléments.',
	de: 'Add-Ins konnten nicht geladen werden.',
	es: 'No se pudieron cargar los complementos.',
	'zh-CN': '无法加载加载项。',
	ja: 'アドインを読み込めませんでした。',
	ko: '추가 기능을 불러오지 못했습니다.',
};
const locales = Object.keys(messages) as (keyof typeof messages)[];

export const addInLoadError = (locale: string) => messages[normalizeLocale(locale, locales, 'en')];
