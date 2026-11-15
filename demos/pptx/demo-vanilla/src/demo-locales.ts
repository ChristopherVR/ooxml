/**
 * Demo-shell string translations (file picker / dropzone, sample-deck loader,
 * and the floating picker tooltips). These keys are not part of the
 * pptx-vanilla-viewer TranslationKey set and are therefore kept in a separate
 * dictionary that is registered alongside each language's viewer bundle in
 * demo-i18n.ts. Text mirrors demos/demo-vue/src/demo-locales.ts, minus
 * the collaboration-join strings (the vanilla binding has no collaboration
 * yet), plus the sample-deck strings the viewer-only demos add.
 */

export const demoStringsEn = {
	'demo.dropzone.hint': 'Drop a .pptx or .ppt file here or click to browse',
	'demo.dropzone.processed': 'The file is processed entirely in the browser',
	'demo.dropzone.browse': 'Browse files',
	'demo.dropzone.newPresentation': 'or create a New Presentation',
	'demo.dropzone.creating': 'Creating...',
	'demo.dropzone.uploadAriaLabel': 'Upload a presentation',
	'demo.viewer.loadError': 'Failed to load the presentation',
	'demo.pickers.switchTheme': 'Switch theme',
	'demo.pickers.switchLanguage': 'Switch language',
	'demo.export.png': 'Export PNG',
	'demo.export.pdf': 'Export PDF',
	'demo.export.gif': 'Export animated GIF',
	'demo.export.video': 'Export WebM video',
	'demo.export.print': 'Print',
};

export const demoStringsFr = {
	'demo.dropzone.hint': 'Deposez un fichier .pptx ou .ppt ici ou cliquez pour parcourir',
	'demo.dropzone.processed': 'Le fichier est traite entierement dans le navigateur',
	'demo.dropzone.browse': 'Parcourir les fichiers',
	'demo.dropzone.newPresentation': 'ou creer une nouvelle presentation',
	'demo.dropzone.creating': 'Creation en cours...',
	'demo.dropzone.uploadAriaLabel': 'Telecharger une presentation',
	'demo.viewer.loadError': 'Echec du chargement de la presentation',
	'demo.pickers.switchTheme': 'Changer de theme',
	'demo.pickers.switchLanguage': 'Changer de langue',
	'demo.export.png': 'Exporter en PNG',
	'demo.export.pdf': 'Exporter en PDF',
	'demo.export.gif': 'Exporter en GIF anime',
	'demo.export.video': 'Exporter en video WebM',
	'demo.export.print': 'Imprimer',
};

export const demoStringsEs = {
	'demo.dropzone.hint': 'Suelte un archivo .pptx o .ppt aqui o haga clic para explorar',
	'demo.dropzone.processed': 'El archivo se procesa completamente en el navegador',
	'demo.dropzone.browse': 'Explorar archivos',
	'demo.dropzone.newPresentation': 'o crear una nueva presentacion',
	'demo.dropzone.creating': 'Creando...',
	'demo.dropzone.uploadAriaLabel': 'Subir una presentacion',
	'demo.viewer.loadError': 'No se pudo cargar la presentacion',
	'demo.pickers.switchTheme': 'Cambiar tema',
	'demo.pickers.switchLanguage': 'Cambiar idioma',
	'demo.export.png': 'Exportar a PNG',
	'demo.export.pdf': 'Exportar a PDF',
	'demo.export.gif': 'Exportar GIF animado',
	'demo.export.video': 'Exportar video WebM',
	'demo.export.print': 'Imprimir',
};

export const demoStringsDe = {
	'demo.dropzone.hint': 'PPTX- oder PPT-Datei hier ablegen oder zum Durchsuchen klicken',
	'demo.dropzone.processed': 'Die Datei wird vollstandig im Browser verarbeitet',
	'demo.dropzone.browse': 'Dateien durchsuchen',
	'demo.dropzone.newPresentation': 'oder eine neue Prasentation erstellen',
	'demo.dropzone.creating': 'Wird erstellt...',
	'demo.dropzone.uploadAriaLabel': 'Praesentation hochladen',
	'demo.viewer.loadError': 'Prasentation konnte nicht geladen werden',
	'demo.pickers.switchTheme': 'Design wechseln',
	'demo.pickers.switchLanguage': 'Sprache wechseln',
	'demo.export.png': 'Als PNG exportieren',
	'demo.export.pdf': 'Als PDF exportieren',
	'demo.export.gif': 'Als animiertes GIF exportieren',
	'demo.export.video': 'Als WebM-Video exportieren',
	'demo.export.print': 'Drucken',
};

export const demoStringsZhCN = {
	'demo.dropzone.hint': '将 .pptx 或 .ppt 文件拖到此处，或单击浏览',
	'demo.dropzone.processed': '文件完全在浏览器中处理',
	'demo.dropzone.browse': '浏览文件',
	'demo.dropzone.newPresentation': '或新建演示文稿',
	'demo.dropzone.creating': '正在创建…',
	'demo.dropzone.uploadAriaLabel': '上传演示文稿',
	'demo.viewer.loadError': '无法加载演示文稿',
	'demo.pickers.switchTheme': '切换主题',
	'demo.pickers.switchLanguage': '切换语言',
	'demo.export.png': '导出 PNG',
	'demo.export.pdf': '导出 PDF',
	'demo.export.gif': '导出 GIF 动画',
	'demo.export.video': '导出视频',
	'demo.export.print': '打印',
} satisfies Record<keyof typeof demoStringsEn, string>;

export const demoStringsJa = {
	'demo.dropzone.hint': '.pptx または .ppt ファイルをここにドロップするか、クリックして参照します',
	'demo.dropzone.processed': 'ファイルはすべてブラウザー内で処理されます',
	'demo.dropzone.browse': 'ファイルを参照',
	'demo.dropzone.newPresentation': 'または新しいプレゼンテーションを作成',
	'demo.dropzone.creating': '作成しています...',
	'demo.dropzone.uploadAriaLabel': 'プレゼンテーションをアップロード',
	'demo.viewer.loadError': 'プレゼンテーションを読み込めませんでした',
	'demo.pickers.switchTheme': 'テーマの切り替え',
	'demo.pickers.switchLanguage': '言語の切り替え',
	'demo.export.png': 'PNG としてエクスポート',
	'demo.export.pdf': 'PDF としてエクスポート',
	'demo.export.gif': 'アニメーション GIF としてエクスポート',
	'demo.export.video': 'WebM ビデオとしてエクスポート',
	'demo.export.print': '印刷',
} satisfies Record<keyof typeof demoStringsEn, string>;

export const demoStringsKo = {
	'demo.dropzone.hint': '.pptx 또는 .ppt 파일을 여기에 끌어 놓거나 클릭하여 찾아보세요',
	'demo.dropzone.processed': '파일은 모두 브라우저에서 처리됩니다',
	'demo.dropzone.browse': '파일 찾아보기',
	'demo.dropzone.newPresentation': '또는 새 프레젠테이션 만들기',
	'demo.dropzone.creating': '만드는 중...',
	'demo.dropzone.uploadAriaLabel': '프레젠테이션 업로드',
	'demo.viewer.loadError': '프레젠테이션을 불러오지 못했습니다',
	'demo.pickers.switchTheme': '테마 전환',
	'demo.pickers.switchLanguage': '언어 전환',
	'demo.export.png': 'PNG로 내보내기',
	'demo.export.pdf': 'PDF로 내보내기',
	'demo.export.gif': '애니메이션 GIF로 내보내기',
	'demo.export.video': 'WebM 비디오로 내보내기',
	'demo.export.print': '인쇄',
} satisfies Record<keyof typeof demoStringsEn, string>;
