import { ShieldAlert, ShieldCheck, ShieldQuestionMark, ShieldX } from '@lucide/vue';
import type { SignatureIcon, SignatureTone } from 'ooxml-ui/pptx';
import type { Component } from 'vue';

/** The Vue icon of each shared signature icon. */
export const SIGNATURE_ICONS: Readonly<Record<SignatureIcon, Component>> = {
	'shield-check': ShieldCheck,
	'shield-alert': ShieldAlert,
	'shield-x': ShieldX,
	'shield-question': ShieldQuestionMark,
};

/** Badge colours (chrome badge and per-row status pill) of each shared tone. */
export const SIGNATURE_BADGE_TONE: Readonly<Record<SignatureTone, string>> = {
	success: 'border-green-700/40 bg-green-900/30 text-green-300 hover:bg-green-900/50',
	caution: 'border-amber-700/40 bg-amber-900/30 text-amber-300 hover:bg-amber-900/50',
	danger: 'border-red-700/40 bg-red-900/30 text-red-300 hover:bg-red-900/50',
	neutral: 'border-border bg-muted/40 text-muted-foreground hover:bg-muted/60',
};

/** Signatures panel header colours of each shared tone. */
export const SIGNATURE_HEADER_TONE: Readonly<Record<SignatureTone, string>> = {
	success: 'bg-green-900/20 text-green-300',
	caution: 'bg-amber-900/20 text-amber-300',
	danger: 'bg-red-900/20 text-red-300',
	neutral: 'bg-muted/30 text-foreground',
};

/** Signatures panel row accent of each shared tone. */
export const SIGNATURE_ROW_TONE: Readonly<Record<SignatureTone, string>> = {
	success: 'border-l-green-500',
	caution: 'border-l-amber-500',
	danger: 'border-l-red-500',
	neutral: 'border-l-border',
};
