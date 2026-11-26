import type { VisioLocale } from '../ribbon-locale';

const en = {
	Shapes: 'Shapes',
	'Close Shapes': 'Close Shapes',
	'Open Shapes': 'Open Shapes',
	'Search shapes': 'Search shapes',
	'Search results': 'Search results',
	'No shapes match your search.': 'No shapes match your search.',
	'More Shapes': 'More Shapes',
	'More Shapes: open another stencil.': 'More Shapes: open another stencil.',
	'Quick Shapes': 'Quick Shapes',
	'Quick Shapes: favourite masters of every open stencil.':
		'Quick Shapes: favourite masters of every open stencil.',
	'No Quick Shapes. Right-click a master to add it.':
		'No Quick Shapes. Right-click a master to add it.',
	'Document Stencil': 'Document Stencil',
	'New Stencil': 'New Stencil',
	'Stencil actions': 'Stencil actions',
	'My Shapes': 'My Shapes',
	'Open Stencil': 'Open Stencil',
	'Open Stencil...': 'Open Stencil...',
	'Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.':
		'Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.',
	'Basic Shapes': 'Basic Shapes',
	'Basic Flowchart Shapes': 'Basic Flowchart Shapes',
	'Miscellaneous Flowchart Shapes': 'Miscellaneous Flowchart Shapes',
	'Arrow Shapes': 'Arrow Shapes',
	Master: 'Master',
	'Add to Quick Shapes': 'Add to Quick Shapes',
	'Remove from Quick Shapes': 'Remove from Quick Shapes',
	'Showing the first {limit} masters; {omitted} more are not listed.':
		'Showing the first {limit} masters; {omitted} more are not listed.',
	'This stencil file is not available here. Its shapes already used in the drawing are in the Document Stencil.':
		'This stencil file is not available here. Its shapes already used in the drawing are in the Document Stencil.',
	'The master has no shapes to drop.': 'The master has no shapes to drop.',
	'A picture or connector master cannot be dropped.':
		'A picture or connector master cannot be dropped.',
	'A master whose shapes are turned cannot be dropped yet.':
		'A master whose shapes are turned cannot be dropped yet.',
	'{name}: drag onto the page, or press Enter to add it at the centre.':
		'{name}: drag onto the page, or press Enter to add it at the centre.',
};

export const VISIO_SHAPES_STRINGS = {
	en,
	de: {
		Shapes: 'Shapes',
		'Close Shapes': 'Shapes schließen',
		'Open Shapes': 'Shapes öffnen',
		'Search shapes': 'Shapes suchen',
		'Search results': 'Suchergebnisse',
		'No shapes match your search.': 'Keine Shapes entsprechen Ihrer Suche.',
		'More Shapes': 'Weitere Shapes',
		'More Shapes: open another stencil.': 'Weitere Shapes: eine andere Schablone öffnen.',
		'Quick Shapes': 'Quick-Shapes',
		'Quick Shapes: favourite masters of every open stencil.':
			'Quick-Shapes: bevorzugte Master-Shapes aller geöffneten Schablonen.',
		'No Quick Shapes. Right-click a master to add it.':
			'Keine Quick-Shapes. Klicken Sie mit der rechten Maustaste auf ein Master-Shape, um es hinzuzufügen.',
		'Document Stencil': 'Dokumentschablone',
		'New Stencil': 'Neue Schablone',
		'Stencil actions': 'Schablonenaktionen',
		'My Shapes': 'Meine Shapes',
		'Open Stencil': 'Schablone öffnen',
		'Open Stencil...': 'Schablone öffnen...',
		'Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.':
			'Schablone öffnen: eine .vssx-Schablonendatei. Ältere .vss-Dateien werden nicht unterstützt.',
		'Basic Shapes': 'Standard-Shapes',
		'Basic Flowchart Shapes': 'Standardflussdiagramm-Shapes',
		'Miscellaneous Flowchart Shapes': 'Sonstige Flussdiagramm-Shapes',
		'Arrow Shapes': 'Pfeil-Shapes',
		Master: 'Master-Shape',
		'Add to Quick Shapes': 'Zu Quick-Shapes hinzufügen',
		'Remove from Quick Shapes': 'Aus Quick-Shapes entfernen',
		'Showing the first {limit} masters; {omitted} more are not listed.':
			'Die ersten {limit} Master-Shapes werden angezeigt. {omitted} weitere sind nicht aufgeführt.',
		'This stencil file is not available here. Its shapes already used in the drawing are in the Document Stencil.':
			'Diese Schablonendatei ist hier nicht verfügbar. Die in der Zeichnung verwendeten Shapes befinden sich in der Dokumentschablone.',
		'The master has no shapes to drop.': 'Das Master-Shape enthält keine Shapes zum Ablegen.',
		'A picture or connector master cannot be dropped.':
			'Ein Bild- oder Verbinder-Master-Shape kann nicht abgelegt werden.',
		'A master whose shapes are turned cannot be dropped yet.':
			'Ein Master-Shape mit gedrehten Shapes kann noch nicht abgelegt werden.',
		'{name}: drag onto the page, or press Enter to add it at the centre.':
			'{name}: Ziehen Sie das Shape auf das Zeichenblatt, oder drücken Sie die EINGABETASTE, um es mittig hinzuzufügen.',
	},
	es: {
		Shapes: 'Formas',
		'Close Shapes': 'Cerrar Formas',
		'Open Shapes': 'Abrir Formas',
		'Search shapes': 'Buscar formas',
		'Search results': 'Resultados de la búsqueda',
		'No shapes match your search.': 'Ninguna forma coincide con la búsqueda.',
		'More Shapes': 'Más formas',
		'More Shapes: open another stencil.': 'Más formas: abrir otra galería de símbolos.',
		'Quick Shapes': 'Formas rápidas',
		'Quick Shapes: favourite masters of every open stencil.':
			'Formas rápidas: formas maestras favoritas de cada galería abierta.',
		'No Quick Shapes. Right-click a master to add it.':
			'No hay formas rápidas. Haga clic con el botón derecho en una forma maestra para agregarla.',
		'Document Stencil': 'Galería de símbolos del documento',
		'New Stencil': 'Nueva galería de símbolos',
		'Stencil actions': 'Acciones de la galería de símbolos',
		'My Shapes': 'Mis formas',
		'Open Stencil': 'Abrir galería de símbolos',
		'Open Stencil...': 'Abrir galería de símbolos...',
		'Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.':
			'Abrir galería de símbolos: un archivo .vssx. No se admiten los archivos .vss antiguos.',
		'Basic Shapes': 'Formas básicas',
		'Basic Flowchart Shapes': 'Formas para diagramas de flujo básicos',
		'Miscellaneous Flowchart Shapes': 'Formas varias para diagramas de flujo',
		'Arrow Shapes': 'Formas de flecha',
		Master: 'Forma maestra',
		'Add to Quick Shapes': 'Agregar a Formas rápidas',
		'Remove from Quick Shapes': 'Quitar de Formas rápidas',
		'Showing the first {limit} masters; {omitted} more are not listed.':
			'Se muestran las primeras {limit} formas maestras. Hay {omitted} más que no se muestran.',
		'This stencil file is not available here. Its shapes already used in the drawing are in the Document Stencil.':
			'Este archivo de galería de símbolos no está disponible aquí. Las formas que ya se utilizan en el dibujo están en la galería de símbolos del documento.',
		'The master has no shapes to drop.': 'La forma maestra no contiene formas para colocar.',
		'A picture or connector master cannot be dropped.':
			'No se puede colocar una forma maestra de imagen o conector.',
		'A master whose shapes are turned cannot be dropped yet.':
			'Aún no se puede colocar una forma maestra con formas giradas.',
		'{name}: drag onto the page, or press Enter to add it at the centre.':
			'{name}: arrastre la forma a la página o presione Entrar para agregarla en el centro.',
	},
	fr: {
		Shapes: 'Formes',
		'Close Shapes': 'Fermer Formes',
		'Open Shapes': 'Ouvrir Formes',
		'Search shapes': 'Rechercher les formes',
		'Search results': 'Résultats de recherche',
		'No shapes match your search.': 'Aucune forme ne correspond à votre recherche.',
		'More Shapes': 'Autres formes',
		'More Shapes: open another stencil.': 'Autres formes : ouvrir un autre gabarit.',
		'Quick Shapes': 'Formes rapides',
		'Quick Shapes: favourite masters of every open stencil.':
			'Formes rapides : formes maîtres favorites de chaque gabarit ouvert.',
		'No Quick Shapes. Right-click a master to add it.':
			'Aucune forme rapide. Cliquez avec le bouton droit sur une forme maître pour l’ajouter.',
		'Document Stencil': 'Gabarit du document',
		'New Stencil': 'Nouveau gabarit',
		'Stencil actions': 'Actions du gabarit',
		'My Shapes': 'Mes formes',
		'Open Stencil': 'Ouvrir le gabarit',
		'Open Stencil...': 'Ouvrir le gabarit...',
		'Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.':
			'Ouvrir le gabarit : un fichier .vssx. Les anciens fichiers .vss ne sont pas pris en charge.',
		'Basic Shapes': 'Formes de base',
		'Basic Flowchart Shapes': 'Formes de diagramme de flux simple',
		'Miscellaneous Flowchart Shapes': 'Formes diverses de diagramme de flux',
		'Arrow Shapes': 'Formes de flèche',
		Master: 'Forme maître',
		'Add to Quick Shapes': 'Ajouter aux formes rapides',
		'Remove from Quick Shapes': 'Retirer des formes rapides',
		'Showing the first {limit} masters; {omitted} more are not listed.':
			'Les {limit} premières formes maîtres sont affichées. {omitted} autres ne sont pas répertoriées.',
		'This stencil file is not available here. Its shapes already used in the drawing are in the Document Stencil.':
			'Ce fichier de gabarit n’est pas disponible ici. Les formes déjà utilisées dans le dessin se trouvent dans le gabarit du document.',
		'The master has no shapes to drop.': 'La forme maître ne contient aucune forme à déposer.',
		'A picture or connector master cannot be dropped.':
			'Une forme maître d’image ou de lien ne peut pas être déposée.',
		'A master whose shapes are turned cannot be dropped yet.':
			'Une forme maître contenant des formes pivotées ne peut pas encore être déposée.',
		'{name}: drag onto the page, or press Enter to add it at the centre.':
			'{name} : faites glisser la forme sur la page ou appuyez sur Entrée pour l’ajouter au centre.',
	},
	'zh-CN': {
		Shapes: '形状',
		'Close Shapes': '关闭形状',
		'Open Shapes': '打开形状',
		'Search shapes': '搜索形状',
		'Search results': '搜索结果',
		'No shapes match your search.': '没有与搜索匹配的形状。',
		'More Shapes': '更多形状',
		'More Shapes: open another stencil.': '更多形状：打开另一个模具。',
		'Quick Shapes': '快速形状',
		'Quick Shapes: favourite masters of every open stencil.':
			'快速形状：所有打开的模具中常用的主控形状。',
		'No Quick Shapes. Right-click a master to add it.': '没有快速形状。右键单击主控形状即可添加。',
		'Document Stencil': '文档模具',
		'New Stencil': '新建模具',
		'Stencil actions': '模具操作',
		'My Shapes': '我的形状',
		'Open Stencil': '打开模具',
		'Open Stencil...': '打开模具...',
		'Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.':
			'打开模具：打开 .vssx 模具文件。不支持旧版 .vss 文件。',
		'Basic Shapes': '基本形状',
		'Basic Flowchart Shapes': '基本流程图形状',
		'Miscellaneous Flowchart Shapes': '其他流程图形状',
		'Arrow Shapes': '箭头形状',
		Master: '主控形状',
		'Add to Quick Shapes': '添加到快速形状',
		'Remove from Quick Shapes': '从快速形状中删除',
		'Showing the first {limit} masters; {omitted} more are not listed.':
			'正在显示前 {limit} 个主控形状，另有 {omitted} 个未列出。',
		'This stencil file is not available here. Its shapes already used in the drawing are in the Document Stencil.':
			'此处无法打开该模具文件。绘图中已使用的形状位于文档模具中。',
		'The master has no shapes to drop.': '主控形状没有可放置的形状。',
		'A picture or connector master cannot be dropped.': '无法放置图片或连接线主控形状。',
		'A master whose shapes are turned cannot be dropped yet.':
			'暂时无法放置包含旋转形状的主控形状。',
		'{name}: drag onto the page, or press Enter to add it at the centre.':
			'{name}：拖动到页面上，或按 Enter 键将其添加到页面中央。',
	},
} satisfies Record<VisioLocale, Record<keyof typeof en, string>>;
