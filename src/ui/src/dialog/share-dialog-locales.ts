/**
 * The shared Share dialog's captions in each display locale the products ship (English, French,
 * German, Spanish, Simplified Chinese), so Word, Excel and Visio need no dictionary entries of
 * their own for it. Pure data; `{count}` is replaced by the number of people.
 */
export interface OfficeShareDialogLabels {
	title: string;
	activeTitle: string;
	description: string;
	createTab: string;
	joinTab: string;
	roomName: string;
	yourName: string;
	serverUrl: string;
	serverPlaceholder: string;
	serverHint: string;
	invitation: string;
	invitationPlaceholder: string;
	start: string;
	join: string;
	cancel: string;
	close: string;
	stop: string;
	link: string;
	copy: string;
	copied: string;
	linkHint: string;
	session: string;
	server: string;
	noServer: string;
	people: string;
	you: string;
	peopleCount: string;
	invalidRoom: string;
	statusOff: string;
	statusConnecting: string;
	statusConnected: string;
	statusJoining: string;
	statusDisconnected: string;
	statusError: string;
}

export const SHARE_DIALOG_LOCALES = ['en', 'fr', 'de', 'es', 'zh-CN'] as const;
export type ShareDialogLocale = (typeof SHARE_DIALOG_LOCALES)[number];

const en: OfficeShareDialogLabels = {
	title: 'Share',
	activeTitle: 'Sharing',
	description:
		'Edit together with other people. Everyone in the same session sees changes as they are made.',
	createTab: 'Start a session',
	joinTab: 'Join a session',
	roomName: 'Session name',
	yourName: 'Your name',
	serverUrl: 'Server (optional)',
	serverPlaceholder: 'wss://collab.example.com',
	serverHint: 'Without a server, only other windows and tabs of this browser can join.',
	invitation: 'Invitation link or session name',
	invitationPlaceholder: 'Paste a link or type a session name',
	start: 'Start sharing',
	join: 'Join',
	cancel: 'Cancel',
	close: 'Close',
	stop: 'Stop sharing',
	link: 'Invitation link',
	copy: 'Copy',
	copied: 'Copied',
	linkHint: 'Send this link to the people you want to edit with.',
	session: 'Session',
	server: 'Server',
	noServer: 'This browser only',
	people: 'People in this session',
	you: '(you)',
	peopleCount: '{count} in this session',
	invalidRoom: 'Use 1 to 128 letters, digits, hyphens or underscores for the session name.',
	statusOff: 'Not sharing.',
	statusConnecting: 'Connecting...',
	statusConnected: 'Connected.',
	statusJoining: 'Joining the session...',
	statusDisconnected: 'Disconnected. Changes are kept and sent when the connection returns.',
	statusError: 'The connection failed. Changes are kept and sent when the connection returns.',
};

const fr: OfficeShareDialogLabels = {
	title: 'Partager',
	activeTitle: 'Partage en cours',
	description:
		'Modifiez à plusieurs. Toutes les personnes de la même session voient les modifications au fur et à mesure.',
	createTab: 'Démarrer une session',
	joinTab: 'Rejoindre une session',
	roomName: 'Nom de la session',
	yourName: 'Votre nom',
	serverUrl: 'Serveur (facultatif)',
	serverPlaceholder: 'wss://collab.example.com',
	serverHint:
		'Sans serveur, seuls les autres onglets et fenêtres de ce navigateur peuvent rejoindre la session.',
	invitation: "Lien d'invitation ou nom de session",
	invitationPlaceholder: 'Collez un lien ou saisissez un nom de session',
	start: 'Commencer le partage',
	join: 'Rejoindre',
	cancel: 'Annuler',
	close: 'Fermer',
	stop: 'Arrêter le partage',
	link: "Lien d'invitation",
	copy: 'Copier',
	copied: 'Copié',
	linkHint: 'Envoyez ce lien aux personnes avec qui vous voulez modifier.',
	session: 'Session',
	server: 'Serveur',
	noServer: 'Ce navigateur uniquement',
	people: 'Personnes dans cette session',
	you: '(vous)',
	peopleCount: '{count} dans cette session',
	invalidRoom:
		'Utilisez de 1 à 128 lettres, chiffres, traits d’union ou traits de soulignement pour le nom de la session.',
	statusOff: 'Aucun partage.',
	statusConnecting: 'Connexion...',
	statusConnected: 'Connecté.',
	statusJoining: 'Connexion à la session...',
	statusDisconnected:
		'Déconnecté. Les modifications sont conservées et envoyées au retour de la connexion.',
	statusError:
		'La connexion a échoué. Les modifications sont conservées et envoyées au retour de la connexion.',
};

const de: OfficeShareDialogLabels = {
	title: 'Freigeben',
	activeTitle: 'Freigegeben',
	description:
		'Gemeinsam mit anderen bearbeiten. Alle in derselben Sitzung sehen Änderungen, sobald sie entstehen.',
	createTab: 'Sitzung starten',
	joinTab: 'Sitzung beitreten',
	roomName: 'Sitzungsname',
	yourName: 'Ihr Name',
	serverUrl: 'Server (optional)',
	serverPlaceholder: 'wss://collab.example.com',
	serverHint: 'Ohne Server können nur andere Fenster und Tabs dieses Browsers beitreten.',
	invitation: 'Einladungslink oder Sitzungsname',
	invitationPlaceholder: 'Link einfügen oder Sitzungsnamen eingeben',
	start: 'Freigabe starten',
	join: 'Beitreten',
	cancel: 'Abbrechen',
	close: 'Schließen',
	stop: 'Freigabe beenden',
	link: 'Einladungslink',
	copy: 'Kopieren',
	copied: 'Kopiert',
	linkHint: 'Senden Sie diesen Link an die Personen, mit denen Sie bearbeiten möchten.',
	session: 'Sitzung',
	server: 'Server',
	noServer: 'Nur dieser Browser',
	people: 'Personen in dieser Sitzung',
	you: '(Sie)',
	peopleCount: '{count} in dieser Sitzung',
	invalidRoom:
		'Verwenden Sie für den Sitzungsnamen 1 bis 128 Buchstaben, Ziffern, Binde- oder Unterstriche.',
	statusOff: 'Keine Freigabe.',
	statusConnecting: 'Verbindung wird hergestellt...',
	statusConnected: 'Verbunden.',
	statusJoining: 'Sitzung wird beigetreten...',
	statusDisconnected:
		'Getrennt. Änderungen bleiben erhalten und werden gesendet, sobald die Verbindung zurück ist.',
	statusError:
		'Die Verbindung ist fehlgeschlagen. Änderungen bleiben erhalten und werden gesendet, sobald die Verbindung zurück ist.',
};

const es: OfficeShareDialogLabels = {
	title: 'Compartir',
	activeTitle: 'Compartiendo',
	description:
		'Edite junto con otras personas. Todos los participantes de la misma sesión ven los cambios a medida que se hacen.',
	createTab: 'Iniciar una sesión',
	joinTab: 'Unirse a una sesión',
	roomName: 'Nombre de la sesión',
	yourName: 'Su nombre',
	serverUrl: 'Servidor (opcional)',
	serverPlaceholder: 'wss://collab.example.com',
	serverHint: 'Sin servidor, solo pueden unirse otras ventanas y pestañas de este explorador.',
	invitation: 'Vínculo de invitación o nombre de sesión',
	invitationPlaceholder: 'Pegue un vínculo o escriba un nombre de sesión',
	start: 'Empezar a compartir',
	join: 'Unirse',
	cancel: 'Cancelar',
	close: 'Cerrar',
	stop: 'Dejar de compartir',
	link: 'Vínculo de invitación',
	copy: 'Copiar',
	copied: 'Copiado',
	linkHint: 'Envíe este vínculo a las personas con las que quiere editar.',
	session: 'Sesión',
	server: 'Servidor',
	noServer: 'Solo este explorador',
	people: 'Personas en esta sesión',
	you: '(usted)',
	peopleCount: '{count} en esta sesión',
	invalidRoom:
		'Use de 1 a 128 letras, dígitos, guiones o guiones bajos para el nombre de la sesión.',
	statusOff: 'Sin compartir.',
	statusConnecting: 'Conectando...',
	statusConnected: 'Conectado.',
	statusJoining: 'Uniéndose a la sesión...',
	statusDisconnected:
		'Desconectado. Los cambios se conservan y se envían cuando vuelva la conexión.',
	statusError: 'Error de conexión. Los cambios se conservan y se envían cuando vuelva la conexión.',
};

const zhCN: OfficeShareDialogLabels = {
	title: '共享',
	activeTitle: '正在共享',
	description: '与他人一起编辑。同一会话中的每个人都会实时看到更改。',
	createTab: '开始会话',
	joinTab: '加入会话',
	roomName: '会话名称',
	yourName: '你的姓名',
	serverUrl: '服务器（可选）',
	serverPlaceholder: 'wss://collab.example.com',
	serverHint: '没有服务器时，只有此浏览器的其他窗口和标签页可以加入。',
	invitation: '邀请链接或会话名称',
	invitationPlaceholder: '粘贴链接或键入会话名称',
	start: '开始共享',
	join: '加入',
	cancel: '取消',
	close: '关闭',
	stop: '停止共享',
	link: '邀请链接',
	copy: '复制',
	copied: '已复制',
	linkHint: '将此链接发送给要一起编辑的人。',
	session: '会话',
	server: '服务器',
	noServer: '仅此浏览器',
	people: '此会话中的人员',
	you: '（你）',
	peopleCount: '此会话中有 {count} 人',
	invalidRoom: '会话名称请使用 1 到 128 个字母、数字、连字符或下划线。',
	statusOff: '未共享。',
	statusConnecting: '正在连接...',
	statusConnected: '已连接。',
	statusJoining: '正在加入会话...',
	statusDisconnected: '已断开连接。更改会保留，并在连接恢复后发送。',
	statusError: '连接失败。更改会保留，并在连接恢复后发送。',
};

/** Captions per locale. */
export const SHARE_DIALOG_LABELS: Readonly<
	Record<ShareDialogLocale, Readonly<OfficeShareDialogLabels>>
> = { en, fr, de, es, 'zh-CN': zhCN };

/** The captions for a product locale (`fr-CA` uses French), English when it is not shipped. */
export function shareDialogLabels(locale: string | undefined): OfficeShareDialogLabels {
	const tag = (locale ?? '').trim();
	if (tag in SHARE_DIALOG_LABELS) return SHARE_DIALOG_LABELS[tag as ShareDialogLocale];
	if (/^zh(-|$)/iu.test(tag)) return zhCN;
	const base = tag.split('-')[0]?.toLowerCase() ?? '';
	return SHARE_DIALOG_LABELS[base as ShareDialogLocale] ?? en;
}
