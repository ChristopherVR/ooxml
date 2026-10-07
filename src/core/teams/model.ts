// Shared types and small validators for the `teams` area: channels, messages, attachments and the
// Office-file kinds a channel can link to. Pure data, no DOM, no Yjs. New code.
import { stripHtmlTags } from '../collab/validation';

export type ChannelKind = 'channel' | 'direct';

export interface Channel {
	id: string;
	name: string;
	topic: string;
	kind: ChannelKind;
	createdAt: number;
	createdBy: string;
	archived: boolean;
}

/** The Office products a channel can link a file from. */
export type OfficeKind = 'docx' | 'xlsx' | 'pptx' | 'vsdx' | 'other';

export interface Attachment {
	name: string;
	kind: OfficeKind;
	size?: number;
	/** Where the file lives: an http(s) URL on the host's storage. Never inlined bytes. */
	url?: string;
	mime?: string;
}

export interface Message {
	id: string;
	channelId: string;
	authorId: string;
	authorName: string;
	text: string;
	ts: number;
	editedAt?: number;
	deleted: boolean;
	/** Id of the thread root this message answers. */
	replyTo?: string;
	attachments: Attachment[];
	/** emoji -> user ids, derived from the reaction map. */
	reactions: Record<string, string[]>;
}

export const MAX_MESSAGE_CHARS = 20_000;
export const MAX_CHANNEL_NAME_CHARS = 80;
export const MAX_ATTACHMENTS = 20;
const MAX_EMOJI_CHARS = 16;

const OFFICE_EXTENSIONS: Record<string, OfficeKind> = {
	docx: 'docx',
	docm: 'docx',
	dotx: 'docx',
	xlsx: 'xlsx',
	xlsm: 'xlsx',
	xltx: 'xlsx',
	pptx: 'pptx',
	pptm: 'pptx',
	potx: 'pptx',
	ppsx: 'pptx',
	vsdx: 'vsdx',
	vsdm: 'vsdx',
	vssx: 'vsdx',
	vstx: 'vsdx',
};

/** Which Office viewer opens a file, from its name (the extension wins over the MIME type). */
export function detectOfficeKind(name: string, mime?: string): OfficeKind {
	const ext = name.split('.').pop()?.toLowerCase() ?? '';
	const byExt = OFFICE_EXTENSIONS[ext];
	if (byExt) return byExt;
	const m = mime?.toLowerCase() ?? '';
	if (m.includes('wordprocessingml')) return 'docx';
	if (m.includes('spreadsheetml')) return 'xlsx';
	if (m.includes('presentationml')) return 'pptx';
	if (m.includes('visio')) return 'vsdx';
	return 'other';
}

/** Message text: tags stripped, control characters removed, trimmed and length-capped. */
export function sanitizeMessageText(text: unknown): string {
	if (typeof text !== 'string') return '';
	const cleaned = stripHtmlTags(text).replace(
		/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/gu,
		'',
	);
	return cleaned.trim().slice(0, MAX_MESSAGE_CHARS);
}

export function sanitizeChannelName(name: unknown): string {
	if (typeof name !== 'string') return '';
	return stripHtmlTags(name).replace(/\s+/gu, ' ').trim().slice(0, MAX_CHANNEL_NAME_CHARS);
}

/** A reaction is one short emoji-like token: no markup, no whitespace. */
export function sanitizeEmoji(emoji: unknown): string {
	if (typeof emoji !== 'string') return '';
	const cleaned = emoji.trim();
	if (!cleaned || [...cleaned].length > MAX_EMOJI_CHARS || /[\s<>&"'|]/u.test(cleaned)) return '';
	return cleaned;
}

/** Attachment links must be http(s) (or a same-origin path); anything else is dropped. */
export function sanitizeAttachment(raw: unknown): Attachment | null {
	if (!raw || typeof raw !== 'object') return null;
	const r = raw as Record<string, unknown>;
	const name = sanitizeChannelName(r.name);
	if (!name) return null;
	const mime = typeof r.mime === 'string' && r.mime.length <= 120 ? r.mime : undefined;
	const out: Attachment = { name, kind: detectOfficeKind(name, mime) };
	if (typeof r.size === 'number' && Number.isFinite(r.size) && r.size >= 0) out.size = r.size;
	if (mime) out.mime = mime;
	if (typeof r.url === 'string' && /^(https?:\/\/|\/(?!\/))/iu.test(r.url) && r.url.length <= 2048)
		out.url = r.url;
	return out;
}

export function sanitizeAttachments(raw: unknown): Attachment[] {
	if (!Array.isArray(raw)) return [];
	const out: Attachment[] = [];
	for (const item of raw.slice(0, MAX_ATTACHMENTS)) {
		const a = sanitizeAttachment(item);
		if (a) out.push(a);
	}
	return out;
}
