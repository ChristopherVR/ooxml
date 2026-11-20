import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureIcon, SignatureStatusView, SignatureTone } from 'ooxml-ui/pptx';
import type { SignatureSummary } from 'pptx-viewer-core';
import { useMemo } from 'react';
import type { IconType } from 'react-icons';
import { LuShieldAlert, LuShieldCheck, LuShieldQuestion, LuShieldX } from 'react-icons/lu';

import { useTranslation } from '../i18n/use-translation';

/**
 * How the React chrome draws the shared signature descriptor (`signatureStatusView`): the
 * icon and colour classes of each icon and tone. The state, wording and tone all come from
 * the descriptor; this file only maps them onto react-icons and Tailwind classes.
 */

export const SIGNATURE_ICONS: Readonly<Record<SignatureIcon, IconType>> = {
	'shield-check': LuShieldCheck,
	'shield-alert': LuShieldAlert,
	'shield-x': LuShieldX,
	'shield-question': LuShieldQuestion,
};

/** A badge or row pill: background, border and text. */
export const SIGNATURE_BADGE_TONE: Readonly<Record<SignatureTone, string>> = {
	success: 'bg-green-900/30 border-green-700/40 text-green-300 hover:bg-green-900/50',
	caution: 'bg-amber-900/30 border-amber-700/40 text-amber-300 hover:bg-amber-900/50',
	danger: 'bg-red-900/30 border-red-700/40 text-red-300 hover:bg-red-900/50',
	neutral: 'bg-muted/30 border-border/40 text-muted-foreground hover:bg-muted/50',
};

/** A banner: a quieter background, border and text. */
export const SIGNATURE_BANNER_TONE: Readonly<Record<SignatureTone, string>> = {
	success: 'bg-green-900/20 border-green-700/30 text-green-200',
	caution: 'bg-amber-900/20 border-amber-700/30 text-amber-200',
	danger: 'bg-red-900/20 border-red-700/30 text-red-200',
	neutral: 'bg-muted/30 border-border/40 text-muted-foreground',
};

/** An icon's colour. */
export const SIGNATURE_ICON_TONE: Readonly<Record<SignatureTone, string>> = {
	success: 'text-green-400',
	caution: 'text-amber-400',
	danger: 'text-red-400',
	neutral: 'text-muted-foreground',
};

/** The deck's signature presentation in the viewer's current UI language. */
export function useSignatureStatusView(summary: SignatureSummary | undefined): SignatureStatusView {
	const { locale } = useTranslation();
	return useMemo(() => signatureStatusView(summary, locale), [summary, locale]);
}
