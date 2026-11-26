import { normalizeLocale } from 'ooxml-core/i18n';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import { ADD_IN_DIALOG_STRINGS } from './add-in-dialog-strings';
import { ADD_IN_MARKETPLACE_STRINGS } from './add-in-marketplace-strings';

const locales = ['fr', 'de', 'es', 'zh-CN', 'ja', 'ko'] as const;
const messages = new Map<string, readonly [string, string, string, string, string, string]>([
	...Object.entries(ADD_IN_DIALOG_STRINGS),
	...Object.entries(ADD_IN_MARKETPLACE_STRINGS),
	[
		'Choose a local add-in file',
		[
			'Choisir un fichier de complément local',
			'Lokale Add-In-Datei auswählen',
			'Elegir un archivo de complemento local',
			'选择本地加载项文件',
			'ローカルのアドイン ファイルを選択',
			'로컬 추가 기능 파일 선택',
		],
	],
	[
		'Installed for this editor. Running add-ins is unavailable.',
		[
			'Installé pour cet éditeur. L’exécution des compléments est indisponible.',
			'Für diesen Editor installiert. Add-Ins können derzeit nicht ausgeführt werden.',
			'Instalado para este editor. No se pueden ejecutar complementos.',
			'已为此编辑器安装。暂时无法运行加载项。',
			'このエディターにインストール済みです。アドインは実行できません。',
			'이 편집기에 설치되었습니다. 추가 기능을 실행할 수 없습니다.',
		],
	],
	[
		'Install for this editor session. Add-in code stays inactive.',
		[
			'Installer pour cette session. Le code du complément reste inactif.',
			'Für diese Editorsitzung installieren. Der Add-In-Code bleibt inaktiv.',
			'Instalar para esta sesión. El código del complemento seguirá inactivo.',
			'为此编辑器会话安装。加载项代码保持停用。',
			'このセッションにインストールします。アドイン コードは実行されません。',
			'이 편집기 세션에 설치합니다. 추가 기능 코드는 비활성 상태로 유지됩니다.',
		],
	],
	[
		'Declared publisher',
		[
			'Éditeur déclaré',
			'Angegebener Herausgeber',
			'Editor declarado',
			'声明的发布者',
			'申告された発行元',
			'명시된 게시자',
		],
	],
	[
		'Origins this add-in may load',
		[
			'Origines que ce complément peut charger',
			'Quellen, von denen dieses Add-In Inhalte laden kann',
			'Orígenes desde los que este complemento puede cargar contenido',
			'此加载项可能从以下来源加载内容',
			'このアドインが読み込む可能性のあるオリジン',
			'이 추가 기능이 콘텐츠를 불러올 수 있는 원본',
		],
	],
	[
		'Review add-in permissions',
		[
			'Vérifier les autorisations du complément',
			'Add-In-Berechtigungen prüfen',
			'Revisar los permisos del complemento',
			'审核加载项权限',
			'アドインのアクセス許可を確認',
			'추가 기능 권한 검토',
		],
	],
	[
		'This add-in requests:',
		[
			'Ce complément demande les autorisations suivantes :',
			'Dieses Add-In fordert Folgendes an:',
			'Este complemento solicita:',
			'此加载项请求：',
			'このアドインは次のアクセス許可を要求します。',
			'이 추가 기능이 다음 권한을 요청합니다.',
		],
	],
	[
		'Use add-in settings and check the active document view',
		[
			'Utiliser les paramètres du complément et consulter l’affichage actif du document',
			'Add-In-Einstellungen verwenden und die aktive Dokumentansicht prüfen',
			'Usar la configuración del complemento y consultar la vista activa del documento',
			'使用加载项设置并查看当前文档视图',
			'アドインの設定を使用し、文書の現在の表示モードを確認する',
			'추가 기능 설정을 사용하고 현재 문서 보기를 확인',
		],
	],
	[
		'Read this document',
		[
			'Lire ce document',
			'Dieses Dokument lesen',
			'Leer este documento',
			'读取此文档',
			'この文書を読み取る',
			'이 문서 읽기',
		],
	],
	[
		'Read and obtain a full copy of this document',
		[
			'Lire ce document et en obtenir une copie complète',
			'Dieses Dokument lesen und eine vollständige Kopie abrufen',
			'Leer y obtener una copia completa de este documento',
			'读取并获取此文档的完整副本',
			'この文書を読み取り、完全なコピーを取得する',
			'이 문서를 읽고 전체 사본 가져오기',
		],
	],
	[
		'Change the selection in this document',
		[
			'Modifier la sélection dans ce document',
			'Die Auswahl in diesem Dokument ändern',
			'Modificar la selección de este documento',
			'更改此文档中选中的内容',
			'この文書で選択されている内容を書き換える',
			'이 문서에서 선택한 내용 변경',
		],
	],
	[
		'Read and change this document',
		[
			'Lire et modifier ce document',
			'Dieses Dokument lesen und ändern',
			'Leer y modificar este documento',
			'读取并更改此文档',
			'この文書を読み取り、変更する',
			'이 문서 읽기 및 변경',
		],
	],
	[
		'Request an unsupported permission ({permission})',
		[
			'Demander une autorisation non prise en charge ({permission})',
			'Eine nicht unterstützte Berechtigung anfordern ({permission})',
			'Solicitar un permiso no compatible ({permission})',
			'请求不受支持的权限（{permission}）',
			'未対応のアクセス許可（{permission}）を要求する',
			'지원되지 않는 권한 요청({permission})',
		],
	],
	[
		'Update this add-in and keep version {version}',
		[
			'Mettre à jour ce complément en conservant la version {version}',
			'Dieses Add-In aktualisieren und Version {version} beibehalten',
			'Actualizar este complemento y mantener la versión {version}',
			'更新此加载项并保留版本 {version}',
			'バージョン {version} のままこのアドインを更新',
			'버전 {version}을 유지하면서 이 추가 기능 업데이트',
		],
	],
	[
		'Update from {oldVersion} to {newVersion}',
		[
			'Mettre à jour de la version {oldVersion} vers la version {newVersion}',
			'Von Version {oldVersion} auf Version {newVersion} aktualisieren',
			'Actualizar de la versión {oldVersion} a la versión {newVersion}',
			'从 {oldVersion} 更新到 {newVersion}',
			'{oldVersion} から {newVersion} に更新',
			'{oldVersion}에서 {newVersion}(으)로 업데이트',
		],
	],
	['None', ['Aucun', 'Keine', 'Ninguno', '无', 'なし', '없음']],
	['Update', ['Mettre à jour', 'Aktualisieren', 'Actualizar', '更新', '更新', '업데이트']],
	[
		'This add-in is already installed.',
		[
			'Ce complément est déjà installé.',
			'Dieses Add-In ist bereits installiert.',
			'Este complemento ya está instalado.',
			'此加载项已安装。',
			'このアドインは既にインストールされています。',
			'이 추가 기능은 이미 설치되었습니다.',
		],
	],
	[
		'Unsupported feature',
		[
			'Fonctionnalité non prise en charge',
			'Nicht unterstützte Funktion',
			'Característica no compatible',
			'不受支持的功能',
			'未対応の機能',
			'지원되지 않는 기능',
		],
	],
	[
		'Unavailable requirement',
		[
			'Configuration requise non satisfaite',
			'Nicht erfüllte Anforderung',
			'Requisito no satisfecho',
			'不满足的要求',
			'満たされていない要件',
			'충족되지 않은 요구 사항',
		],
	],
	[
		'The add-in manifest is invalid.',
		[
			'Le manifeste du complément est invalide.',
			'Das Add-In-Manifest ist ungültig.',
			'El manifiesto del complemento no es válido.',
			'加载项清单无效。',
			'アドイン マニフェストが無効です。',
			'추가 기능 매니페스트가 잘못되었습니다.',
		],
	],
	[
		'The add-in package is invalid or missing a file.',
		[
			'Le package du complément est invalide ou il manque un fichier.',
			'Das Add-In-Paket ist ungültig oder eine Datei fehlt.',
			'El paquete del complemento no es válido o le falta un archivo.',
			'加载项包无效或缺少文件。',
			'アドイン パッケージが無効か、ファイルが不足しています。',
			'추가 기능 패키지가 잘못되었거나 파일이 누락되었습니다.',
		],
	],
	[
		'The add-in file exceeds the size or entry limit.',
		[
			'Le fichier du complément dépasse la limite de taille ou de nombre d’entrées.',
			'Die Add-In-Datei überschreitet die Grenze für Größe oder Anzahl der Einträge.',
			'El archivo supera el límite de tamaño o de entradas.',
			'加载项文件超过大小或条目数限制。',
			'アドイン ファイルがサイズまたは項目数の上限を超えています。',
			'추가 기능 파일이 크기 또는 항목 수 제한을 초과합니다.',
		],
	],
	[
		'The add-in specifies an unsafe URL or origin.',
		[
			'Le complément indique une URL ou une origine non sûre.',
			'Das Add-In gibt eine unsichere URL oder einen unsicheren Ursprung an.',
			'El complemento especifica una URL o un origen no seguros.',
			'加载项指定了不安全的 URL 或来源。',
			'アドインに安全でない URL またはオリジンが指定されています。',
			'추가 기능에 안전하지 않은 URL 또는 원본이 지정되었습니다.',
		],
	],
	[
		'The add-in does not support this editor.',
		[
			'Le complément ne prend pas en charge cet éditeur.',
			'Das Add-In unterstützt diesen Editor nicht.',
			'El complemento no es compatible con este editor.',
			'此加载项不支持此编辑器。',
			'このアドインはこのエディターに対応していません。',
			'이 추가 기능은 이 편집기를 지원하지 않습니다.',
		],
	],
	[
		'This manifest version is unsupported.',
		[
			'Cette version du manifeste n’est pas prise en charge.',
			'Diese Manifestversion wird nicht unterstützt.',
			'Esta versión del manifiesto no es compatible.',
			'不支持此清单版本。',
			'このマニフェスト バージョンには対応していません。',
			'이 매니페스트 버전은 지원되지 않습니다.',
		],
	],
	[
		'This add-in type is unsupported.',
		[
			'Ce type de complément n’est pas pris en charge.',
			'Dieser Add-In-Typ wird nicht unterstützt.',
			'Este tipo de complemento no es compatible.',
			'不支持此加载项类型。',
			'この種類のアドインには対応していません。',
			'이 추가 기능 유형은 지원되지 않습니다.',
		],
	],
]);

export function addInLocale(locale: string): string {
	return normalizeLocale(locale, locales, 'en');
}

export function localAddInText(locale: string, fallback: OfficeServiceText): OfficeServiceText {
	const index = locales.findIndex((candidate) => candidate === locale);
	return (label) => messages.get(label)?.[index] ?? fallback(label);
}

const permissionLabels = new Map([
	['Document.Restricted.User', 'Use add-in settings and check the active document view'],
	['Document.Read.User', 'Read this document'],
	['Document.ReadAll.User', 'Read and obtain a full copy of this document'],
	['Document.Write.User', 'Change the selection in this document'],
	['Document.ReadWrite.User', 'Read and change this document'],
]);

export function permissionText(permission: string, t: OfficeServiceText): string {
	const label = permissionLabels.get(permission);
	return label
		? t(label)
		: t('Request an unsupported permission ({permission})').replace('{permission}', permission);
}
