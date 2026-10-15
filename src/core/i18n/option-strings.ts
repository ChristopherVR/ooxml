/**
 * Display strings of the File > Options dialog that every product shares (category names, the
 * dialog's buttons, the common sections and the reasons a setting is not available). Keys are the
 * English text, so a product merges them into its own dictionary. `{product}` is replaced with
 * the product name (Word, Excel, Visio) by the options builders in `ooxml-ui`.
 */
export type OptionStrings = Record<string, string>;

export const enOptionStrings: OptionStrings = {
	General: 'General',
	Display: 'Display',
	Formulas: 'Formulas',
	Data: 'Data',
	Proofing: 'Proofing',
	Save: 'Save',
	Language: 'Language',
	Accessibility: 'Accessibility',
	Advanced: 'Advanced',
	'Customize Ribbon': 'Customize Ribbon',
	'Quick Access Toolbar': 'Quick Access Toolbar',
	'Add-ins': 'Add-ins',
	'Trust Center': 'Trust Center',
	OK: 'OK',
	Cancel: 'Cancel',
	Close: 'Close',
	Categories: 'Categories',
	'Reset all options': 'Reset all options',
	'General options for working with {product}.': 'General options for working with {product}.',
	'Change how {product} corrects and formats your text.':
		'Change how {product} corrects and formats your text.',
	'Customize how documents are saved.': 'Customize how documents are saved.',
	'Set the Office Language Preferences.': 'Set the Office Language Preferences.',
	'Make {product} easier to work with.': 'Make {product} easier to work with.',
	'Advanced options for working with {product}.': 'Advanced options for working with {product}.',
	'Customize the Ribbon and keyboard shortcuts.': 'Customize the Ribbon and keyboard shortcuts.',
	'Customize the Quick Access Toolbar.': 'Customize the Quick Access Toolbar.',
	'View and manage Microsoft Office Add-ins.': 'View and manage Microsoft Office Add-ins.',
	'Help keep your documents safe and your computer secure and healthy.':
		'Help keep your documents safe and your computer secure and healthy.',
	'User Interface options': 'User Interface options',
	'Show Mini Toolbar on selection': 'Show Mini Toolbar on selection',
	'Enable Live Preview': 'Enable Live Preview',
	'ScreenTip style': 'ScreenTip style',
	'Show feature descriptions in ScreenTips': 'Show feature descriptions in ScreenTips',
	'Personalize your copy of Microsoft Office': 'Personalize your copy of Microsoft Office',
	'User name': 'User name',
	'Office Theme': 'Office Theme',
	'Office display language': 'Office display language',
	'Choose the language for menus, buttons and messages.':
		'Choose the language for menus, buttons and messages.',
	'Not available yet. Needs rich text formatting.':
		'Not available yet. Needs rich text formatting.',
	'Not available yet. Styles are applied on click; hover previews are not drawn yet.':
		'Not available yet. Styles are applied on click; hover previews are not drawn yet.',
	'Not available yet. ScreenTips always describe the command.':
		'Not available yet. ScreenTips always describe the command.',
	'Not available yet. Documents are saved to this device; AutoRecover is not kept.':
		'Not available yet. Documents are saved to this device; AutoRecover is not kept.',
	'Not available yet. There are no accessibility settings to change.':
		'Not available yet. There are no accessibility settings to change.',
	'Not available yet. The editing behaviour is fixed.':
		'Not available yet. The editing behaviour is fixed.',
	'Not available yet. The Quick Access Toolbar is fixed.':
		'Not available yet. The Quick Access Toolbar is fixed.',
	'Not available yet. The ribbon layout is fixed.':
		'Not available yet. The ribbon layout is fixed.',
	'Not available yet. The display language is fixed.':
		'Not available yet. The display language is fixed.',
	'Not available. Add-ins and macros are never run.':
		'Not available. Add-ins and macros are never run.',
	'Not available. Documents stay local and active content is never run.':
		'Not available. Documents stay local and active content is never run.',
};

const frOptionStrings: OptionStrings = {
	General: 'Général',
	Display: 'Affichage',
	Formulas: 'Formules',
	Data: 'Données',
	Proofing: 'Vérification',
	Save: 'Enregistrer',
	Language: 'Langue',
	Accessibility: 'Accessibilité',
	Advanced: 'Options avancées',
	'Customize Ribbon': 'Personnaliser le ruban',
	'Quick Access Toolbar': 'Barre d’outils Accès rapide',
	'Add-ins': 'Compléments',
	'Trust Center': 'Centre de gestion de la confidentialité',
	OK: 'OK',
	Cancel: 'Annuler',
	Close: 'Fermer',
	Categories: 'Catégories',
	'Reset all options': 'Réinitialiser toutes les options',
	'General options for working with {product}.':
		'Options générales pour l’utilisation de {product}.',
	'Change how {product} corrects and formats your text.':
		'Modifiez la façon dont {product} corrige et met en forme votre texte.',
	'Customize how documents are saved.': 'Personnalisez l’enregistrement des documents.',
	'Set the Office Language Preferences.': 'Définissez les préférences de langue d’Office.',
	'Make {product} easier to work with.': 'Facilitez l’utilisation de {product}.',
	'Advanced options for working with {product}.':
		'Options avancées pour l’utilisation de {product}.',
	'Customize the Ribbon and keyboard shortcuts.':
		'Personnalisez le ruban et les raccourcis clavier.',
	'Customize the Quick Access Toolbar.': 'Personnalisez la barre d’outils Accès rapide.',
	'View and manage Microsoft Office Add-ins.':
		'Affichez et gérez les compléments Microsoft Office.',
	'Help keep your documents safe and your computer secure and healthy.':
		'Protégez vos documents et assurez la sécurité et le bon état de votre ordinateur.',
	'User Interface options': 'Options de l’interface utilisateur',
	'Show Mini Toolbar on selection': 'Afficher la mini barre d’outils lors de la sélection',
	'Enable Live Preview': 'Activer l’aperçu dynamique',
	'ScreenTip style': 'Style des info-bulles',
	'Show feature descriptions in ScreenTips':
		'Afficher les descriptions des fonctionnalités dans les info-bulles',
	'Personalize your copy of Microsoft Office': 'Personnalisez votre copie de Microsoft Office',
	'User name': 'Nom d’utilisateur',
	'Office Theme': 'Thème Office',
	'Office display language': 'Langue d’affichage d’Office',
	'Choose the language for menus, buttons and messages.':
		'Choisissez la langue des menus, des boutons et des messages.',
	'Not available yet. Needs rich text formatting.':
		'Pas encore disponible. Nécessite la mise en forme de texte enrichi.',
	'Not available yet. Styles are applied on click; hover previews are not drawn yet.':
		'Pas encore disponible. Les styles s’appliquent au clic ; les aperçus au survol ne sont pas encore dessinés.',
	'Not available yet. ScreenTips always describe the command.':
		'Pas encore disponible. Les info-bulles décrivent toujours la commande.',
	'Not available yet. Documents are saved to this device; AutoRecover is not kept.':
		'Pas encore disponible. Les documents sont enregistrés sur cet appareil ; la récupération automatique n’est pas conservée.',
	'Not available yet. There are no accessibility settings to change.':
		'Pas encore disponible. Aucun paramètre d’accessibilité n’est modifiable.',
	'Not available yet. The editing behaviour is fixed.':
		'Pas encore disponible. Le comportement de modification est fixe.',
	'Not available yet. The Quick Access Toolbar is fixed.':
		'Pas encore disponible. La barre d’outils Accès rapide est fixe.',
	'Not available yet. The ribbon layout is fixed.':
		'Pas encore disponible. La disposition du ruban est fixe.',
	'Not available yet. The display language is fixed.':
		'Pas encore disponible. La langue d’affichage est fixe.',
	'Not available. Add-ins and macros are never run.':
		'Non disponible. Les compléments et les macros ne sont jamais exécutés.',
	'Not available. Documents stay local and active content is never run.':
		'Non disponible. Les documents restent locaux et le contenu actif n’est jamais exécuté.',
};

const deOptionStrings: OptionStrings = {
	General: 'Allgemein',
	Display: 'Anzeige',
	Formulas: 'Formeln',
	Data: 'Daten',
	Proofing: 'Dokumentprüfung',
	Save: 'Speichern',
	Language: 'Sprache',
	Accessibility: 'Barrierefreiheit',
	Advanced: 'Erweitert',
	'Customize Ribbon': 'Menüband anpassen',
	'Quick Access Toolbar': 'Symbolleiste für den Schnellzugriff',
	'Add-ins': 'Add-Ins',
	'Trust Center': 'Trust Center',
	OK: 'OK',
	Cancel: 'Abbrechen',
	Close: 'Schließen',
	Categories: 'Kategorien',
	'Reset all options': 'Alle Optionen zurücksetzen',
	'General options for working with {product}.':
		'Allgemeine Optionen für die Arbeit mit {product}.',
	'Change how {product} corrects and formats your text.':
		'Ändern Sie, wie {product} Ihren Text korrigiert und formatiert.',
	'Customize how documents are saved.': 'Passen Sie an, wie Dokumente gespeichert werden.',
	'Set the Office Language Preferences.': 'Legen Sie die Office-Spracheinstellungen fest.',
	'Make {product} easier to work with.': 'Machen Sie {product} einfacher bedienbar.',
	'Advanced options for working with {product}.':
		'Erweiterte Optionen für die Arbeit mit {product}.',
	'Customize the Ribbon and keyboard shortcuts.':
		'Passen Sie das Menüband und die Tastenkombinationen an.',
	'Customize the Quick Access Toolbar.': 'Passen Sie die Symbolleiste für den Schnellzugriff an.',
	'View and manage Microsoft Office Add-ins.': 'Microsoft Office-Add-Ins anzeigen und verwalten.',
	'Help keep your documents safe and your computer secure and healthy.':
		'Schützen Sie Ihre Dokumente und halten Sie Ihren Computer sicher und fehlerfrei.',
	'User Interface options': 'Benutzeroberflächenoptionen',
	'Show Mini Toolbar on selection': 'Minisymbolleiste bei Auswahl anzeigen',
	'Enable Live Preview': 'Livevorschau aktivieren',
	'ScreenTip style': 'QuickInfo-Format',
	'Show feature descriptions in ScreenTips': 'Funktionsbeschreibungen in QuickInfos anzeigen',
	'Personalize your copy of Microsoft Office': 'Microsoft Office personalisieren',
	'User name': 'Benutzername',
	'Office Theme': 'Office-Design',
	'Office display language': 'Office-Anzeigesprache',
	'Choose the language for menus, buttons and messages.':
		'Wählen Sie die Sprache für Menüs, Schaltflächen und Meldungen.',
	'Not available yet. Needs rich text formatting.':
		'Noch nicht verfügbar. Erfordert Formatierung von Rich-Text.',
	'Not available yet. Styles are applied on click; hover previews are not drawn yet.':
		'Noch nicht verfügbar. Formatvorlagen werden per Klick angewendet; Vorschauen beim Darüberfahren werden noch nicht gezeichnet.',
	'Not available yet. ScreenTips always describe the command.':
		'Noch nicht verfügbar. QuickInfos beschreiben immer den Befehl.',
	'Not available yet. Documents are saved to this device; AutoRecover is not kept.':
		'Noch nicht verfügbar. Dokumente werden auf diesem Gerät gespeichert; AutoWiederherstellen wird nicht geführt.',
	'Not available yet. There are no accessibility settings to change.':
		'Noch nicht verfügbar. Es gibt keine Einstellungen zur Barrierefreiheit zu ändern.',
	'Not available yet. The editing behaviour is fixed.':
		'Noch nicht verfügbar. Das Bearbeitungsverhalten ist festgelegt.',
	'Not available yet. The Quick Access Toolbar is fixed.':
		'Noch nicht verfügbar. Die Symbolleiste für den Schnellzugriff ist festgelegt.',
	'Not available yet. The ribbon layout is fixed.':
		'Noch nicht verfügbar. Das Menüband-Layout ist festgelegt.',
	'Not available yet. The display language is fixed.':
		'Noch nicht verfügbar. Die Anzeigesprache ist festgelegt.',
	'Not available. Add-ins and macros are never run.':
		'Nicht verfügbar. Add-Ins und Makros werden nie ausgeführt.',
	'Not available. Documents stay local and active content is never run.':
		'Nicht verfügbar. Dokumente bleiben lokal und aktive Inhalte werden nie ausgeführt.',
};

const esOptionStrings: OptionStrings = {
	General: 'General',
	Display: 'Mostrar',
	Formulas: 'Fórmulas',
	Data: 'Datos',
	Proofing: 'Revisión',
	Save: 'Guardar',
	Language: 'Idioma',
	Accessibility: 'Accesibilidad',
	Advanced: 'Avanzadas',
	'Customize Ribbon': 'Personalizar cinta de opciones',
	'Quick Access Toolbar': 'Barra de herramientas de acceso rápido',
	'Add-ins': 'Complementos',
	'Trust Center': 'Centro de confianza',
	OK: 'Aceptar',
	Cancel: 'Cancelar',
	Close: 'Cerrar',
	Categories: 'Categorías',
	'Reset all options': 'Restablecer todas las opciones',
	'General options for working with {product}.': 'Opciones generales para trabajar con {product}.',
	'Change how {product} corrects and formats your text.':
		'Cambie la forma en que {product} corrige y da formato al texto.',
	'Customize how documents are saved.': 'Personalice cómo se guardan los documentos.',
	'Set the Office Language Preferences.': 'Establezca las preferencias de idioma de Office.',
	'Make {product} easier to work with.': 'Haga que {product} sea más fácil de usar.',
	'Advanced options for working with {product}.': 'Opciones avanzadas para trabajar con {product}.',
	'Customize the Ribbon and keyboard shortcuts.':
		'Personalice la cinta de opciones y los métodos abreviados de teclado.',
	'Customize the Quick Access Toolbar.': 'Personalice la barra de herramientas de acceso rápido.',
	'View and manage Microsoft Office Add-ins.':
		'Vea y administre los complementos de Microsoft Office.',
	'Help keep your documents safe and your computer secure and healthy.':
		'Ayude a mantener seguros sus documentos y a proteger y mantener en buen estado su equipo.',
	'User Interface options': 'Opciones de la interfaz de usuario',
	'Show Mini Toolbar on selection': 'Mostrar minibarra de herramientas al seleccionar',
	'Enable Live Preview': 'Habilitar vistas previas dinámicas',
	'ScreenTip style': 'Estilo de la información en pantalla',
	'Show feature descriptions in ScreenTips':
		'Mostrar descripciones de características en la información en pantalla',
	'Personalize your copy of Microsoft Office': 'Personalice su copia de Microsoft Office',
	'User name': 'Nombre de usuario',
	'Office Theme': 'Tema de Office',
	'Office display language': 'Idioma de presentación de Office',
	'Choose the language for menus, buttons and messages.':
		'Elija el idioma de los menús, los botones y los mensajes.',
	'Not available yet. Needs rich text formatting.':
		'Aún no disponible. Requiere formato de texto enriquecido.',
	'Not available yet. Styles are applied on click; hover previews are not drawn yet.':
		'Aún no disponible. Los estilos se aplican al hacer clic; las vistas previas al pasar el cursor todavía no se dibujan.',
	'Not available yet. ScreenTips always describe the command.':
		'Aún no disponible. La información en pantalla siempre describe el comando.',
	'Not available yet. Documents are saved to this device; AutoRecover is not kept.':
		'Aún no disponible. Los documentos se guardan en este dispositivo; no se conserva la recuperación automática.',
	'Not available yet. There are no accessibility settings to change.':
		'Aún no disponible. No hay opciones de accesibilidad que cambiar.',
	'Not available yet. The editing behaviour is fixed.':
		'Aún no disponible. El comportamiento de edición es fijo.',
	'Not available yet. The Quick Access Toolbar is fixed.':
		'Aún no disponible. La barra de herramientas de acceso rápido es fija.',
	'Not available yet. The ribbon layout is fixed.':
		'Aún no disponible. El diseño de la cinta de opciones es fijo.',
	'Not available yet. The display language is fixed.':
		'Aún no disponible. El idioma de presentación es fijo.',
	'Not available. Add-ins and macros are never run.':
		'No disponible. Los complementos y las macros nunca se ejecutan.',
	'Not available. Documents stay local and active content is never run.':
		'No disponible. Los documentos permanecen locales y el contenido activo nunca se ejecuta.',
};

const zhOptionStrings: OptionStrings = {
	General: '常规',
	Display: '显示',
	Formulas: '公式',
	Data: '数据',
	Proofing: '校对',
	Save: '保存',
	Language: '语言',
	Accessibility: '辅助功能',
	Advanced: '高级',
	'Customize Ribbon': '自定义功能区',
	'Quick Access Toolbar': '快速访问工具栏',
	'Add-ins': '加载项',
	'Trust Center': '信任中心',
	OK: '确定',
	Cancel: '取消',
	Close: '关闭',
	Categories: '类别',
	'Reset all options': '重置所有选项',
	'General options for working with {product}.': '使用 {product} 的常规选项。',
	'Change how {product} corrects and formats your text.':
		'更改 {product} 更正和设置文本格式的方式。',
	'Customize how documents are saved.': '自定义保存文档的方式。',
	'Set the Office Language Preferences.': '设置 Office 语言首选项。',
	'Make {product} easier to work with.': '让 {product} 更易于使用。',
	'Advanced options for working with {product}.': '使用 {product} 的高级选项。',
	'Customize the Ribbon and keyboard shortcuts.': '自定义功能区和键盘快捷方式。',
	'Customize the Quick Access Toolbar.': '自定义快速访问工具栏。',
	'View and manage Microsoft Office Add-ins.': '查看和管理 Microsoft Office 加载项。',
	'Help keep your documents safe and your computer secure and healthy.':
		'帮助保障文档安全，并使计算机保持安全和健康。',
	'User Interface options': '用户界面选项',
	'Show Mini Toolbar on selection': '选择时显示浮动工具栏',
	'Enable Live Preview': '启用实时预览',
	'ScreenTip style': '屏幕提示样式',
	'Show feature descriptions in ScreenTips': '在屏幕提示中显示功能说明',
	'Personalize your copy of Microsoft Office': '对 Microsoft Office 进行个性化设置',
	'User name': '用户名',
	'Office Theme': 'Office 主题',
	'Office display language': 'Office 显示语言',
	'Choose the language for menus, buttons and messages.': '选择菜单、按钮和消息使用的语言。',
	'Not available yet. Needs rich text formatting.': '暂不可用。需要富文本格式。',
	'Not available yet. Styles are applied on click; hover previews are not drawn yet.':
		'暂不可用。样式在单击时应用；尚不绘制悬停预览。',
	'Not available yet. ScreenTips always describe the command.':
		'暂不可用。屏幕提示始终描述该命令。',
	'Not available yet. Documents are saved to this device; AutoRecover is not kept.':
		'暂不可用。文档保存到此设备；不保留自动恢复。',
	'Not available yet. There are no accessibility settings to change.':
		'暂不可用。没有可更改的辅助功能设置。',
	'Not available yet. The editing behaviour is fixed.': '暂不可用。编辑行为是固定的。',
	'Not available yet. The Quick Access Toolbar is fixed.': '暂不可用。快速访问工具栏是固定的。',
	'Not available yet. The ribbon layout is fixed.': '暂不可用。功能区布局是固定的。',
	'Not available yet. The display language is fixed.': '暂不可用。显示语言是固定的。',
	'Not available. Add-ins and macros are never run.': '不可用。加载项和宏永远不会运行。',
	'Not available. Documents stay local and active content is never run.':
		'不可用。文档保留在本地，活动内容永远不会运行。',
};

/** The shared options strings of one display locale (`en`, `fr`, `de`, `es`, `zh-CN`). */
export const optionStrings = {
	en: enOptionStrings,
	fr: frOptionStrings,
	de: deOptionStrings,
	es: esOptionStrings,
	'zh-CN': zhOptionStrings,
} as const satisfies Record<string, OptionStrings>;
