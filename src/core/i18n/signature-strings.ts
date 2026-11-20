/**
 * The words every product uses for a document's digital signature state (`SignatureState` of
 * `ooxml-core/opc`): a short label for a badge or status bar entry, a sentence for an info panel
 * or load notice, and the note that saving an edit removes the signatures. One table, so Word,
 * Excel, PowerPoint and Visio never describe the same state differently.
 */
import type { SignatureState } from '../opc/signature/state';
import { normalizeLocale } from './locale';

/** The text of one signed state. */
export interface SignatureStateText {
	/** A few words for a badge or status bar entry. */
	label: string;
	/** One or two sentences for an info panel or a load notice. */
	description: string;
}

/** A locale's signature words. `unsigned` has no text: nothing is shown for it. */
export interface SignatureStrings {
	states: Record<Exclude<SignatureState, 'unsigned'>, SignatureStateText>;
	/** Appended to a notice: why an edited, saved copy is no longer signed. */
	removedOnSave: string;
	/** The heading of a list of signatures, and its signer and signing-time captions. */
	heading: string;
	signer: string;
	signedAt: string;
	/** The load notice's buttons: open the list of signatures, and close the notice. */
	view: string;
	dismiss: string;
}

const en: SignatureStrings = {
	states: {
		verified: {
			label: 'Signed',
			description:
				'This document is digitally signed. The signature is valid and the signer is trusted.',
		},
		'verified-untrusted': {
			label: 'Signed, signer not verified',
			description:
				'This document is digitally signed and has not changed since it was signed, but the signer’s certificate is not trusted here, so who signed it is not confirmed.',
		},
		invalid: {
			label: 'Invalid signature',
			description:
				'This document has a digital signature that is not valid. The document may have been changed since it was signed.',
		},
		unsupported: {
			label: 'Signature not verified',
			description:
				'This document has a digital signature that could not be verified here, so it is not shown as signed.',
		},
	},
	removedOnSave: 'Saving changes removes the signatures, because any change invalidates them.',
	heading: 'Digital signatures',
	signer: 'Signer',
	signedAt: 'Signed on',
	view: 'View signatures',
	dismiss: 'Dismiss',
};

const fr: SignatureStrings = {
	states: {
		verified: {
			label: 'Signé',
			description:
				'Ce document est signé numériquement. La signature est valide et le signataire est approuvé.',
		},
		'verified-untrusted': {
			label: 'Signé, signataire non vérifié',
			description:
				'Ce document est signé numériquement et n’a pas été modifié depuis sa signature, mais le certificat du signataire n’est pas approuvé ici : l’identité du signataire n’est donc pas confirmée.',
		},
		invalid: {
			label: 'Signature non valide',
			description:
				'Ce document contient une signature numérique non valide. Il a peut-être été modifié depuis sa signature.',
		},
		unsupported: {
			label: 'Signature non vérifiée',
			description:
				'Ce document contient une signature numérique qui n’a pas pu être vérifiée ici : il n’est donc pas présenté comme signé.',
		},
	},
	removedOnSave:
		'L’enregistrement de modifications supprime les signatures, car toute modification les invalide.',
	heading: 'Signatures numériques',
	signer: 'Signataire',
	signedAt: 'Signé le',
	view: 'Afficher les signatures',
	dismiss: 'Ignorer',
};

const de: SignatureStrings = {
	states: {
		verified: {
			label: 'Signiert',
			description:
				'Dieses Dokument ist digital signiert. Die Signatur ist gültig, und der Signierer ist vertrauenswürdig.',
		},
		'verified-untrusted': {
			label: 'Signiert, Signierer nicht überprüft',
			description:
				'Dieses Dokument ist digital signiert und wurde seit dem Signieren nicht geändert, aber dem Zertifikat des Signierers wird hier nicht vertraut. Wer es signiert hat, ist daher nicht bestätigt.',
		},
		invalid: {
			label: 'Ungültige Signatur',
			description:
				'Dieses Dokument hat eine ungültige digitale Signatur. Das Dokument wurde möglicherweise seit dem Signieren geändert.',
		},
		unsupported: {
			label: 'Signatur nicht überprüft',
			description:
				'Dieses Dokument hat eine digitale Signatur, die hier nicht überprüft werden konnte. Es wird daher nicht als signiert angezeigt.',
		},
	},
	removedOnSave:
		'Beim Speichern von Änderungen werden die Signaturen entfernt, da jede Änderung sie ungültig macht.',
	heading: 'Digitale Signaturen',
	signer: 'Signierer',
	signedAt: 'Signiert am',
	view: 'Signaturen anzeigen',
	dismiss: 'Schließen',
};

const es: SignatureStrings = {
	states: {
		verified: {
			label: 'Firmado',
			description:
				'Este documento está firmado digitalmente. La firma es válida y el firmante es de confianza.',
		},
		'verified-untrusted': {
			label: 'Firmado, firmante no comprobado',
			description:
				'Este documento está firmado digitalmente y no ha cambiado desde que se firmó, pero el certificado del firmante no es de confianza aquí, por lo que no se ha confirmado quién lo firmó.',
		},
		invalid: {
			label: 'Firma no válida',
			description:
				'Este documento tiene una firma digital que no es válida. Es posible que el documento haya cambiado desde que se firmó.',
		},
		unsupported: {
			label: 'Firma no comprobada',
			description:
				'Este documento tiene una firma digital que no se ha podido comprobar aquí, por lo que no se muestra como firmado.',
		},
	},
	removedOnSave: 'Al guardar cambios se quitan las firmas, porque cualquier cambio las invalida.',
	heading: 'Firmas digitales',
	signer: 'Firmante',
	signedAt: 'Firmado el',
	view: 'Ver firmas',
	dismiss: 'Descartar',
};

const zh: SignatureStrings = {
	states: {
		verified: {
			label: '已签名',
			description: '此文档已进行数字签名。签名有效，且签名者受信任。',
		},
		'verified-untrusted': {
			label: '已签名，签名者未验证',
			description:
				'此文档已进行数字签名，且签名后未被更改，但签名者的证书在此处不受信任，因此无法确认签名者的身份。',
		},
		invalid: {
			label: '签名无效',
			description: '此文档的数字签名无效。文档在签名后可能已被更改。',
		},
		unsupported: {
			label: '签名未验证',
			description: '此文档的数字签名无法在此处验证，因此不会显示为已签名。',
		},
	},
	removedOnSave: '保存更改会删除签名，因为任何更改都会使签名失效。',
	heading: '数字签名',
	signer: '签名者',
	signedAt: '签名时间',
	view: '查看签名',
	dismiss: '关闭',
};

const ja: SignatureStrings = {
	states: {
		verified: {
			label: '署名済み',
			description:
				'このドキュメントにはデジタル署名があります。署名は有効で、署名者は信頼されています。',
		},
		'verified-untrusted': {
			label: '署名済み (署名者は未確認)',
			description:
				'このドキュメントにはデジタル署名があり、署名後に変更されていません。ただし、署名者の証明書はここでは信頼されていないため、署名者は確認されていません。',
		},
		invalid: {
			label: '無効な署名',
			description:
				'このドキュメントのデジタル署名は無効です。署名後にドキュメントが変更された可能性があります。',
		},
		unsupported: {
			label: '署名は未検証',
			description:
				'このドキュメントのデジタル署名はここでは検証できなかったため、署名済みとして表示されません。',
		},
	},
	removedOnSave: '変更を保存すると署名は削除されます。変更によって署名は無効になるためです。',
	heading: 'デジタル署名',
	signer: '署名者',
	signedAt: '署名日時',
	view: '署名の表示',
	dismiss: '閉じる',
};

const ko: SignatureStrings = {
	states: {
		verified: {
			label: '서명됨',
			description: '이 문서는 디지털 서명되었습니다. 서명이 유효하며 서명자를 신뢰할 수 있습니다.',
		},
		'verified-untrusted': {
			label: '서명됨, 서명자 확인 안 됨',
			description:
				'이 문서는 디지털 서명되었으며 서명 후 변경되지 않았지만, 여기에서는 서명자의 인증서를 신뢰할 수 없으므로 서명한 사람이 확인되지 않았습니다.',
		},
		invalid: {
			label: '잘못된 서명',
			description:
				'이 문서의 디지털 서명이 유효하지 않습니다. 서명 후 문서가 변경되었을 수 있습니다.',
		},
		unsupported: {
			label: '서명 확인 안 됨',
			description:
				'이 문서의 디지털 서명을 여기에서 확인할 수 없으므로 서명된 문서로 표시되지 않습니다.',
		},
	},
	removedOnSave: '변경 내용을 저장하면 서명이 제거됩니다. 변경하면 서명이 무효화되기 때문입니다.',
	heading: '디지털 서명',
	signer: '서명자',
	signedAt: '서명 날짜',
	view: '서명 보기',
	dismiss: '닫기',
};

/** The signature words of each display locale. */
export const signatureStrings = {
	en,
	fr,
	de,
	es,
	'zh-CN': zh,
	ja,
	ko,
} as const satisfies Record<string, SignatureStrings>;

const SIGNATURE_LOCALES = Object.keys(signatureStrings) as (keyof typeof signatureStrings)[];

/**
 * The signature words for `locale`, resolved the way the product UIs resolve theirs
 * (`normalizeLocale`); an unsupported tag reads English.
 */
export const signatureStringsFor = (locale: string | undefined): SignatureStrings =>
	signatureStrings[normalizeLocale(locale, SIGNATURE_LOCALES, 'en')];

/** The label and description of `state` in `locale`, or `undefined` for an unsigned document. */
export function signatureStateText(
	state: SignatureState,
	locale: string | undefined,
): SignatureStateText | undefined {
	return state === 'unsigned' ? undefined : signatureStringsFor(locale).states[state];
}
