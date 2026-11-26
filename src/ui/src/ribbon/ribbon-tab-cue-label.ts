const cueLabels: Readonly<Record<string, string>> = {
	fr: 'Autres onglets',
	de: 'Weitere Registerkarten',
	es: 'Más pestañas',
	'zh-CN': '更多选项卡',
	ja: 'その他のタブ',
	ko: '탭 더 보기',
};

export const ribbonTabCueLabel = (locale: string) => cueLabels[locale] ?? 'More tabs';
