import { normalizeLocale } from 'ooxml-core/i18n';
import type { OfficeServiceText } from '../dialog/office-service-dialog';

const locales = ['fr', 'de', 'es', 'zh-CN', 'ja', 'ko'] as const;
const messages = new Map<string, readonly [string, string, string, string, string, string]>([
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
		'Declared origins',
		[
			'Origines déclarées',
			'Angegebene Ursprünge',
			'Orígenes declarados',
			'声明的来源',
			'申告されたオリジン',
			'명시된 원본',
		],
	],
	[
		'Requested permissions',
		[
			'Autorisations demandées',
			'Angeforderte Berechtigungen',
			'Permisos solicitados',
			'请求的权限',
			'要求された権限',
			'요청된 권한',
		],
	],
	[
		'Installed version',
		[
			'Version installée',
			'Installierte Version',
			'Versión instalada',
			'已安装的版本',
			'インストール済みのバージョン',
			'설치된 버전',
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

export function localAddInText(locale: string, fallback: OfficeServiceText): OfficeServiceText {
	const language = normalizeLocale(locale, locales, 'en');
	const index = locales.findIndex((candidate) => candidate === language);
	return (label) => messages.get(label)?.[index] ?? fallback(label);
}
