import {
	LucideShieldCheck,
	LucideShieldExclamationPoint,
	LucideShieldQuestionMark,
	LucideShieldX,
} from '@lucide/angular';
import type { LucideIconInput } from '@lucide/angular';
import type { SignatureIcon } from 'ooxml-ui/pptx';

/** The Lucide shield of each signature state icon `signatureStatusView` names. */
export const SIGNATURE_ICONS: Readonly<Record<SignatureIcon, LucideIconInput>> = {
	'shield-check': LucideShieldCheck,
	'shield-alert': LucideShieldExclamationPoint,
	'shield-x': LucideShieldX,
	'shield-question': LucideShieldQuestionMark,
};
