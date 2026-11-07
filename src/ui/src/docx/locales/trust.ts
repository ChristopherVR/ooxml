/** File > Options > Trust Center: PowerPoint's wording, with Word named in the info text. */
export const enTrust = {
	'Help keep your documents safe and secure.': 'Help keep your documents safe and secure.',
	'Security settings': 'Security settings',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'These settings help keep the content you open safe. We recommend the defaults.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Download fonts that are not installed or embedded (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.':
		'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.',
	'This setting is managed by your organization': 'This setting is managed by your organization',
};

type Trust = Record<keyof typeof enTrust, string>;

export const frTrust: Trust = {
	'Help keep your documents safe and secure.':
		'Assurez la sécurité et la protection de vos documents.',
	'Security settings': 'Paramètres de sécurité',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'Ces paramètres contribuent à la sécurité du contenu que vous ouvrez. Nous recommandons les valeurs par défaut.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Télécharger les polices non installées ni incorporées (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.':
		"Lorsque cette option est activée, votre navigateur contacte fonts.googleapis.com et fonts.gstatic.com (ou le serveur de polices de votre organisation) pour obtenir les polices manquantes. Lorsqu'elle est désactivée, le texte utilise des polices installées, incorporées ou de substitution, et les sauts de ligne peuvent donc différer de Word.",
	'This setting is managed by your organization': 'Ce paramètre est géré par votre organisation',
};

export const deTrust: Trust = {
	'Help keep your documents safe and secure.': 'Sorgen Sie für die Sicherheit Ihrer Dokumente.',
	'Security settings': 'Sicherheitseinstellungen',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'Diese Einstellungen tragen dazu bei, geöffnete Inhalte sicher zu halten. Wir empfehlen die Standardwerte.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Nicht installierte oder eingebettete Schriftarten herunterladen (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.':
		'Wenn aktiviert, ruft Ihr Browser fehlende Schriftarten von fonts.googleapis.com und fonts.gstatic.com (oder vom Schriftartserver Ihrer Organisation) ab. Wenn deaktiviert, verwendet Text installierte, eingebettete oder Ersatzschriftarten, sodass Zeilenumbrüche von Word abweichen können.',
	'This setting is managed by your organization':
		'Diese Einstellung wird von Ihrer Organisation verwaltet',
};

export const esTrust: Trust = {
	'Help keep your documents safe and secure.':
		'Ayude a mantener sus documentos seguros y protegidos.',
	'Security settings': 'Configuración de seguridad',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'Esta configuración ayuda a mantener seguro el contenido que abre. Recomendamos los valores predeterminados.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Descargar fuentes que no están instaladas ni incrustadas (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.':
		'Cuando está activado, el navegador se conecta a fonts.googleapis.com y fonts.gstatic.com (o al servidor de fuentes de su organización) para obtener las fuentes que faltan. Cuando está desactivado, el texto usa fuentes instaladas, incrustadas o sustitutas, por lo que los saltos de línea pueden diferir de Word.',
	'This setting is managed by your organization':
		'Esta configuración la administra su organización',
};

export const zhTrust: Trust = {
	'Help keep your documents safe and secure.': '帮助确保您的文档安全可靠。',
	'Security settings': '安全设置',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'这些设置有助于确保您打开的内容安全。我们建议使用默认值。',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'下载未安装或未嵌入的字体（Google Fonts）',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Word.':
		'开启时，浏览器会连接 fonts.googleapis.com 和 fonts.gstatic.com（或贵组织设置的字体服务器）以获取缺少的字体。关闭时，文本使用已安装、已嵌入或替代字体，因此换行可能与 Word 不同。',
	'This setting is managed by your organization': '此设置由你的组织管理',
};
