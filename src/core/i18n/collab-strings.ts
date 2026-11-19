/**
 * Display strings every product shows when a collaboration session refuses something: an update
 * that stopped sharing (`UpdateRefusedError` of `ooxml-core/collab`) or values a collaborator sent
 * that were left out. Keys are the English text, the `message` of the error, so a product merges
 * them into its own dictionary and translates the message it was given.
 */
import { LEFT_OUT_MESSAGE, UPDATE_REFUSAL_MESSAGES } from '../collab/refusal';
import { normalizeLocale } from './locale';

export type CollabStrings = Record<string, string>;

const { 'too-large': tooLarge, malformed, 'server-rejected': rejected } = UPDATE_REFUSAL_MESSAGES;
const full = UPDATE_REFUSAL_MESSAGES['server-full'];

const en: CollabStrings = Object.fromEntries(
	[tooLarge, malformed, rejected, full, LEFT_OUT_MESSAGE].map((text) => [text, text]),
);

const fr: CollabStrings = {
	[tooLarge]:
		"Une modification de la session partagée est trop volumineuse pour être appliquée ici. Le partage s'est donc arrêté. Votre copie n'a pas été modifiée.",
	[malformed]:
		"Une modification de la session partagée n'a pas pu être lue. Le partage s'est donc arrêté. Votre copie n'a pas été modifiée.",
	[rejected]:
		"Le serveur a refusé votre dernière modification. Le partage s'est donc arrêté. Votre copie conserve la modification.",
	[full]:
		"Le document partagé a atteint la taille maximale du serveur. Le partage s'est donc arrêté. Votre copie conserve votre dernière modification.",
	[LEFT_OUT_MESSAGE]:
		"Du contenu envoyé par un collaborateur a été ignoré, car il n'est pas autorisé ici, par exemple un lien ou un type de fichier non sécurisé.",
};

const de: CollabStrings = {
	[tooLarge]:
		'Eine Änderung aus der freigegebenen Sitzung ist zu groß, um hier übernommen zu werden, daher wurde die Freigabe beendet. Ihre Kopie ist unverändert.',
	[malformed]:
		'Eine Änderung aus der freigegebenen Sitzung konnte nicht gelesen werden, daher wurde die Freigabe beendet. Ihre Kopie ist unverändert.',
	[rejected]:
		'Der Server hat Ihre letzte Änderung abgelehnt, daher wurde die Freigabe beendet. Ihre Kopie behält die Änderung.',
	[full]:
		'Das freigegebene Dokument hat die Größenbeschränkung des Servers erreicht, daher wurde die Freigabe beendet. Ihre Kopie behält Ihre letzte Änderung.',
	[LEFT_OUT_MESSAGE]:
		'Inhalte, die ein Mitbearbeiter gesendet hat, wurden ausgelassen, weil sie hier nicht zulässig sind, etwa ein unsicherer Link oder Dateityp.',
};

const es: CollabStrings = {
	[tooLarge]:
		'Un cambio de la sesión compartida es demasiado grande para aplicarlo aquí, por lo que se detuvo el uso compartido. Su copia no ha cambiado.',
	[malformed]:
		'No se pudo leer un cambio de la sesión compartida, por lo que se detuvo el uso compartido. Su copia no ha cambiado.',
	[rejected]:
		'El servidor rechazó su último cambio, por lo que se detuvo el uso compartido. Su copia conserva el cambio.',
	[full]:
		'El documento compartido alcanzó el límite de tamaño del servidor, por lo que se detuvo el uso compartido. Su copia conserva su último cambio.',
	[LEFT_OUT_MESSAGE]:
		'Se omitió contenido enviado por un colaborador porque no está permitido aquí, como un vínculo o un tipo de archivo no seguro.',
};

const zh: CollabStrings = {
	[tooLarge]: '共享会话中的一项更改太大，无法在此应用，因此共享已停止。你的副本未更改。',
	[malformed]: '无法读取共享会话中的一项更改，因此共享已停止。你的副本未更改。',
	[rejected]: '服务器拒绝了你的最新更改，因此共享已停止。你的副本保留了该更改。',
	[full]: '共享文档已达到服务器的大小限制，因此共享已停止。你的副本保留了你的最新更改。',
	[LEFT_OUT_MESSAGE]: '协作者发送的内容因不允许在此使用（例如不安全的链接或文件类型）而被省略。',
};

const ja: CollabStrings = {
	[tooLarge]:
		'共有セッションからの変更が大きすぎてここでは適用できないため、共有を停止しました。お使いのコピーは変更されていません。',
	[malformed]:
		'共有セッションからの変更を読み取れなかったため、共有を停止しました。お使いのコピーは変更されていません。',
	[rejected]:
		'サーバーが最新の変更を拒否したため、共有を停止しました。お使いのコピーにはその変更が残っています。',
	[full]:
		'共有ドキュメントがサーバーのサイズ上限に達したため、共有を停止しました。お使いのコピーには最新の変更が残っています。',
	[LEFT_OUT_MESSAGE]:
		'共同編集者が送信したコンテンツは、安全でないリンクやファイルの種類など、ここでは許可されていないため除外されました。',
};

const ko: CollabStrings = {
	[tooLarge]:
		'공유 세션의 변경 내용이 너무 커서 여기에 적용할 수 없으므로 공유가 중지되었습니다. 사용자의 사본은 변경되지 않았습니다.',
	[malformed]:
		'공유 세션의 변경 내용을 읽을 수 없으므로 공유가 중지되었습니다. 사용자의 사본은 변경되지 않았습니다.',
	[rejected]:
		'서버가 최근 변경 내용을 거부했으므로 공유가 중지되었습니다. 사용자의 사본에는 변경 내용이 남아 있습니다.',
	[full]:
		'공유 문서가 서버의 크기 제한에 도달했으므로 공유가 중지되었습니다. 사용자의 사본에는 최근 변경 내용이 남아 있습니다.',
	[LEFT_OUT_MESSAGE]:
		'공동 작업자가 보낸 콘텐츠는 안전하지 않은 링크나 파일 형식처럼 여기에서 허용되지 않아 제외되었습니다.',
};

/** The shared collaboration strings of each display locale. */
export const collabStrings = {
	en,
	fr,
	de,
	es,
	'zh-CN': zh,
	ja,
	ko,
} as const satisfies Record<string, CollabStrings>;

const COLLAB_LOCALES = Object.keys(collabStrings) as (keyof typeof collabStrings)[];

/**
 * `text` in `locale` when it is one of these strings, else `text`. Any BCP 47 tag resolves the
 * way the product UIs resolve theirs (`normalizeLocale`): `zh`, `zh-Hans` and `zh_CN` read
 * Simplified Chinese, `de-AT` reads German, and `zh-TW` or an unsupported tag reads English.
 */
export function translateCollabText(text: string, locale: string | undefined): string {
	const table: CollabStrings = collabStrings[normalizeLocale(locale, COLLAB_LOCALES, 'en')];
	return table[text] ?? text;
}
