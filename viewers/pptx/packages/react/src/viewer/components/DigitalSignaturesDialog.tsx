import type { SignatureStatusRow } from 'ooxml-ui/pptx';
import type { SignatureSummary } from 'pptx-viewer-core';
import React from 'react';
import { LuInfo, LuX } from 'react-icons/lu';

import { useTranslation } from '../i18n/use-translation';
import { DialogFooter } from './DialogFooter';
import {
	SIGNATURE_BADGE_TONE,
	SIGNATURE_BANNER_TONE,
	SIGNATURE_ICONS,
	SIGNATURE_ICON_TONE,
	useSignatureStatusView,
} from './signature-look';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/**
 * Props for the {@link DigitalSignaturesDialog} component.
 */
export interface DigitalSignaturesDialogProps {
	/** Whether the dialog is visible. */
	isOpen: boolean;
	/** Callback invoked when the dialog is dismissed. */
	onClose: () => void;
	/** The loaded deck's signatures (`PptxData.signatures`), absent when it has none. */
	signatures: SignatureSummary | undefined;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

function SignatureRow({
	row,
	signedAtCaption,
}: {
	row: SignatureStatusRow;
	signedAtCaption: string;
}): React.ReactElement {
	const { t } = useTranslation();
	const Icon = SIGNATURE_ICONS[row.icon];
	return (
		<li
			className='rounded-lg border border-border/40 px-3 py-2 space-y-1'
			data-pptx-signature-row-state={row.state}
		>
			<div className='flex items-center justify-between gap-2'>
				<span className='text-xs font-semibold text-foreground break-words' title={row.subject}>
					{row.signer}
				</span>
				<span
					className={`inline-flex shrink-0 items-center gap-1 whitespace-nowrap rounded-full border px-2 py-0.5 text-[11px] font-semibold ${SIGNATURE_BADGE_TONE[row.tone]}`}
				>
					<Icon className='w-3 h-3' aria-hidden='true' />
					{row.label}
				</span>
			</div>
			{(row.issuer || row.signedAt) && (
				<dl className='grid grid-cols-[auto_1fr] gap-x-2.5 gap-y-0.5 text-[11px] text-muted-foreground'>
					{row.issuer && (
						<>
							<dt className='font-medium'>{t('pptx.digitalSignatures.issuer')}</dt>
							<dd className='m-0 break-words'>{row.issuer}</dd>
						</>
					)}
					{row.signedAt && (
						<>
							<dt className='font-medium'>{signedAtCaption}</dt>
							<dd className='m-0'>{row.signedAt}</dd>
						</>
					)}
				</dl>
			)}
		</li>
	);
}

/**
 * Modal dialog that displays the digital signature state of a PPTX file.
 *
 * When signatures are present it shows the verified state (a deck reads "Signed" only when its
 * signatures verify), each signature with its own state, and the note that saving removes them.
 * When no signatures are found, an informational notice is displayed instead.
 *
 * @param props - {@link DigitalSignaturesDialogProps}
 * @returns The dialog element, or `null` when `isOpen` is `false`.
 */
export function DigitalSignaturesDialog({
	isOpen,
	onClose,
	signatures,
}: DigitalSignaturesDialogProps): React.ReactElement | null {
	const { t } = useTranslation();
	const view = useSignatureStatusView(signatures);

	if (!isOpen) {
		return null;
	}

	const Icon = SIGNATURE_ICONS[view.icon];
	return (
		<>
			{/* Backdrop */}
			<button
				type='button'
				style={{ zIndex: 1200 }}
				className='fixed inset-0 bg-black/60'
				aria-label={t('pptx.common.close')}
				onClick={onClose}
			/>
			{/* Dialog */}
			<div
				style={{ zIndex: 1201 }}
				className='fixed inset-0 flex items-center justify-center pointer-events-none'
			>
				<div
					className='pointer-events-auto w-[420px] rounded-xl border border-border bg-popover backdrop-blur-xl shadow-2xl max-md:fixed max-md:inset-x-0 max-md:bottom-0 max-md:top-auto max-md:w-full max-md:max-h-[88dvh] max-md:overflow-y-auto max-md:rounded-t-2xl max-md:rounded-b-none max-md:border-x-0 max-md:border-b-0 max-md:pb-[max(env(safe-area-inset-bottom),0px)]'
					data-pptx-signature-state={view.state}
				>
					{/* Header */}
					<div className='flex items-center justify-between px-5 py-4 border-b border-border/60'>
						<div className='flex items-center gap-2'>
							<Icon className={`w-5 h-5 ${SIGNATURE_ICON_TONE[view.tone]}`} aria-hidden='true' />
							<h2 className='text-sm font-semibold text-foreground'>{view.heading}</h2>
						</div>
						<button
							type='button'
							onClick={onClose}
							className='p-1 rounded hover:bg-accent transition-colors'
							aria-label={t('pptx.common.close')}
						>
							<LuX className='w-4 h-4 text-muted-foreground' />
						</button>
					</div>

					{/* Body */}
					<div className='px-5 py-6 space-y-4'>
						{view.present ? (
							<>
								{/* Overall state */}
								<div
									className={`flex items-start gap-3 rounded-lg border px-4 py-3 ${SIGNATURE_BANNER_TONE[view.tone]}`}
								>
									<Icon
										className={`w-5 h-5 shrink-0 mt-0.5 ${SIGNATURE_ICON_TONE[view.tone]}`}
										aria-hidden='true'
									/>
									<div className='space-y-1'>
										<p className='text-xs font-semibold'>{view.label}</p>
										<p className='text-[11px] opacity-80'>{view.description}</p>
										<p className='text-[11px] opacity-70'>
											{t('pptx.digitalSignatures.signatureCount', { count: view.count })}
										</p>
									</div>
								</div>

								{/* Each signature */}
								{view.rows.length > 0 && (
									<ul className='space-y-2'>
										{view.rows.map((row) => (
											<SignatureRow
												key={row.path}
												row={row}
												signedAtCaption={view.signedAtCaption}
											/>
										))}
									</ul>
								)}

								{/* Saving removes them */}
								<div className='flex items-start gap-3 rounded-lg bg-amber-900/20 border border-amber-700/30 px-4 py-3'>
									<LuInfo className='w-5 h-5 text-amber-400 shrink-0 mt-0.5' />
									<p className='text-xs text-amber-200'>{view.removedOnSave}</p>
								</div>
							</>
						) : (
							<>
								{/* No signatures */}
								<div className='flex items-start gap-3 rounded-lg bg-muted/30 border border-border/40 px-4 py-3'>
									<LuInfo className='w-5 h-5 text-muted-foreground shrink-0 mt-0.5' />
									<p className='text-xs text-muted-foreground'>
										{t('pptx.digitalSignatures.noSignatures')}
									</p>
								</div>
							</>
						)}
					</div>

					{/* Footer */}
					<div className='flex items-center justify-end px-5 py-3 border-t border-border/60'>
						<DialogFooter
							actions={[{ id: 'close', label: t('pptx.common.close') }]}
							onAction={onClose}
						/>
					</div>
				</div>
			</div>
		</>
	);
}
