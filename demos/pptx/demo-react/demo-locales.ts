/**
 * Demo-shell string translations (file picker / dropzone and collaboration
 * join messages). These keys are not part of the pptx-react-viewer TranslationKey
 * set and are therefore kept in a separate dictionary that is merged into each
 * language's resource bundle in i18n.ts.
 */

export const demoStringsEn = {
	'demo.dropzone.hint': 'Drop a .pptx or .ppt file here or click to browse',
	'demo.dropzone.hintCollab':
		'Drop a .pptx or .ppt file here or click to browse to start collaborating',
	'demo.dropzone.processed': 'The file is processed entirely in the browser',
	'demo.dropzone.browse': 'Browse files',
	'demo.dropzone.newPresentation': 'or create a New Presentation',
	'demo.dropzone.creating': 'Creating...',
	'demo.dropzone.joiningBroadcast': 'Joining broadcast:',
	'demo.dropzone.loadingBroadcast': 'Loading presentation from broadcaster\u2026',
	'demo.dropzone.joiningSession': 'Joining collaboration session:',
	'demo.dropzone.uploadAriaLabel': 'Upload a presentation',
};

export const demoStringsFr = {
	'demo.dropzone.hint': 'Deposez un fichier .pptx ou .ppt ici ou cliquez pour parcourir',
	'demo.dropzone.hintCollab':
		'Deposez un fichier .pptx ou .ppt ici ou cliquez pour demarrer la collaboration',
	'demo.dropzone.processed': 'Le fichier est traite entierement dans le navigateur',
	'demo.dropzone.browse': 'Parcourir les fichiers',
	'demo.dropzone.newPresentation': 'ou creer une nouvelle presentation',
	'demo.dropzone.creating': 'Creation...',
	'demo.dropzone.joiningBroadcast': 'Rejoindre la diffusion :',
	'demo.dropzone.loadingBroadcast': 'Chargement de la presentation depuis le diffuseur...',
	'demo.dropzone.joiningSession': 'Rejoindre la session de collaboration :',
	'demo.dropzone.uploadAriaLabel': 'Telecharger une presentation',
};

export const demoStringsEs = {
	'demo.dropzone.hint': 'Suelte un archivo .pptx o .ppt aqui o haga clic para explorar',
	'demo.dropzone.hintCollab':
		'Suelte un archivo .pptx o .ppt aqui o haga clic para comenzar a colaborar',
	'demo.dropzone.processed': 'El archivo se procesa completamente en el navegador',
	'demo.dropzone.browse': 'Explorar archivos',
	'demo.dropzone.newPresentation': 'o crear una nueva presentacion',
	'demo.dropzone.creating': 'Creando...',
	'demo.dropzone.joiningBroadcast': 'Uniendose a la transmision:',
	'demo.dropzone.loadingBroadcast': 'Cargando presentacion desde el emisor...',
	'demo.dropzone.joiningSession': 'Uniendose a la sesion de colaboracion:',
	'demo.dropzone.uploadAriaLabel': 'Subir una presentacion',
};

export const demoStringsDe = {
	'demo.dropzone.hint': 'PPTX- oder PPT-Datei hier ablegen oder zum Durchsuchen klicken',
	'demo.dropzone.hintCollab':
		'PPTX- oder PPT-Datei hier ablegen oder klicken, um die Zusammenarbeit zu starten',
	'demo.dropzone.processed': 'Die Datei wird vollstandig im Browser verarbeitet',
	'demo.dropzone.browse': 'Dateien durchsuchen',
	'demo.dropzone.newPresentation': 'oder eine neue Prasentation erstellen',
	'demo.dropzone.creating': 'Wird erstellt...',
	'demo.dropzone.joiningBroadcast': 'Ubertragung beitreten:',
	'demo.dropzone.loadingBroadcast': 'Prasentation vom Sender wird geladen...',
	'demo.dropzone.joiningSession': 'Zusammenarbeitssitzung beitreten:',
	'demo.dropzone.uploadAriaLabel': 'Praesentation hochladen',
};

export const demoStringsZhCN = {
	'demo.dropzone.hint': '将 .pptx 或 .ppt 文件拖到此处，或单击浏览',
	'demo.dropzone.hintCollab': '将 .pptx 或 .ppt 文件拖到此处，或单击浏览以开始协作',
	'demo.dropzone.processed': '文件完全在浏览器中处理',
	'demo.dropzone.browse': '浏览文件',
	'demo.dropzone.newPresentation': '或新建演示文稿',
	'demo.dropzone.creating': '正在创建…',
	'demo.dropzone.joiningBroadcast': '正在加入广播：',
	'demo.dropzone.loadingBroadcast': '正在从广播者加载演示文稿…',
	'demo.dropzone.joiningSession': '正在加入协作会话：',
	'demo.dropzone.uploadAriaLabel': '上传演示文稿',
} satisfies Record<keyof typeof demoStringsEn, string>;

export const demoStringsJa = {
	'demo.dropzone.hint': '.pptx または .ppt ファイルをここにドロップするか、クリックして参照します',
	'demo.dropzone.hintCollab':
		'.pptx または .ppt ファイルをここにドロップするか、クリックして参照し、共同編集を開始します',
	'demo.dropzone.processed': 'ファイルはすべてブラウザー内で処理されます',
	'demo.dropzone.browse': 'ファイルを参照',
	'demo.dropzone.newPresentation': 'または新しいプレゼンテーションを作成',
	'demo.dropzone.creating': '作成しています...',
	'demo.dropzone.joiningBroadcast': 'ブロードキャストに参加しています:',
	'demo.dropzone.loadingBroadcast': 'ブロードキャスト元からプレゼンテーションを読み込んでいます...',
	'demo.dropzone.joiningSession': '共同編集セッションに参加しています:',
	'demo.dropzone.uploadAriaLabel': 'プレゼンテーションをアップロード',
} satisfies Record<keyof typeof demoStringsEn, string>;

export const demoStringsKo = {
	'demo.dropzone.hint': '.pptx 또는 .ppt 파일을 여기에 끌어 놓거나 클릭하여 찾아보세요',
	'demo.dropzone.hintCollab':
		'.pptx 또는 .ppt 파일을 여기에 끌어 놓거나 클릭하여 찾아보고 공동 작업을 시작하세요',
	'demo.dropzone.processed': '파일은 모두 브라우저에서 처리됩니다',
	'demo.dropzone.browse': '파일 찾아보기',
	'demo.dropzone.newPresentation': '또는 새 프레젠테이션 만들기',
	'demo.dropzone.creating': '만드는 중...',
	'demo.dropzone.joiningBroadcast': '브로드캐스트에 참가하는 중:',
	'demo.dropzone.loadingBroadcast': '브로드캐스터에서 프레젠테이션을 불러오는 중...',
	'demo.dropzone.joiningSession': '공동 작업 세션에 참가하는 중:',
	'demo.dropzone.uploadAriaLabel': '프레젠테이션 업로드',
} satisfies Record<keyof typeof demoStringsEn, string>;
