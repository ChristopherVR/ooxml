import type {
	BuildingBlockResult,
	HostSignatureInfo,
	ViewerBuildingBlocksResult,
} from 'pptx-react-viewer';
import React, { useEffect, useState } from 'react';

/**
 * Host-owned UI for the file-level dialogs the building blocks never mount: password protection,
 * document properties, font embedding and digital signatures. The toolbar controls raise a request
 * (`actions.pending`); this panel reads the group's state and applies through its operations. Every
 * result is written to a readout so the browser spec can assert state, outcome and failure codes.
 */
function describe(result: BuildingBlockResult): string {
	return result.ok ? 'ok' : result.code;
}

export function HostOwnedFileDialogs({
	blocks,
}: {
	blocks: ViewerBuildingBlocksResult;
}): React.ReactElement {
	const { actions, toolbarProps } = blocks;
	const [last, setLast] = useState('none');
	const [password, setPassword] = useState('');
	const [confirmation, setConfirmation] = useState('');
	const [title, setTitle] = useState('');
	const [signatures, setSignatures] = useState<HostSignatureInfo[] | null>(null);
	const report = (name: string, result: BuildingBlockResult) =>
		setLast(`${name}: ${describe(result)}`);
	const pending = actions.pending?.kind;
	const protection = actions.passwordProtection.state;
	const properties = actions.documentProperties.current;
	const fonts = actions.fontEmbedding.state;
	const signed = actions.digitalSignatures.state;
	const { list } = actions.digitalSignatures;
	useEffect(() => {
		if (pending === 'digitalSignatures') {
			void list().then(setSignatures);
		}
	}, [pending, list]);
	return (
		<aside aria-label='Host file dialogs' data-host-file-dialogs style={{ padding: 4 }}>
			<div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
				<button type='button' onClick={() => toolbarProps.onOpenPasswordProtection?.()}>
					Host protection
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenDocumentProperties?.()}>
					Host properties
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenFontEmbedding?.()}>
					Host fonts
				</button>
				<button type='button' onClick={() => toolbarProps.onOpenDigitalSignatures?.()}>
					Host signatures
				</button>
			</div>
			<output aria-label='Host file dialog result'>{last}</output>
			<output aria-label='Host protection state'>
				{`open=${protection.openPassword} modify=${protection.modifyPassword} readOnly=${protection.readOnlyRecommended} final=${protection.markedFinal} modifySupported=${protection.supported.modifyPassword}`}
			</output>
			<output aria-label='Host properties state'>
				{`title=${properties.core.title ?? ''} custom=${properties.custom.map((p) => `${p.name}:${p.value}`).join(',')}`}
			</output>
			<output aria-label='Host fonts state'>
				{`embed=${fonts.embedFonts} canEmbed=${fonts.canEmbed} embedded=${fonts.embeddedFamilies.length} used=${fonts.usedFamilies.length}`}
			</output>
			<output aria-label='Host signatures state'>
				{`signed=${signed.signed} count=${signed.count} listed=${signatures?.length ?? 'loading'} create=${signed.supported.create}`}
			</output>
			{pending === 'passwordProtection' && (
				<div role='group' aria-label='Host protection form'>
					<input
						type='password'
						aria-label='Host password'
						value={password}
						onChange={(e) => setPassword(e.target.value)}
					/>
					<input
						type='password'
						aria-label='Host password confirmation'
						value={confirmation}
						onChange={(e) => setConfirmation(e.target.value)}
					/>
					<button
						type='button'
						onClick={() =>
							report('password', actions.passwordProtection.setOpenPassword(password, confirmation))
						}
					>
						Set password
					</button>
					<button
						type='button'
						onClick={() => report('password', actions.passwordProtection.clearOpenPassword())}
					>
						Clear password
					</button>
					<button
						type='button'
						onClick={() => report('final', actions.passwordProtection.setMarkedFinal(true))}
					>
						Mark as final
					</button>
					<button
						type='button'
						onClick={() => report('modify', actions.passwordProtection.setModifyPassword('x'))}
					>
						Set modify password
					</button>
				</div>
			)}
			{pending === 'documentProperties' && (
				<div role='group' aria-label='Host properties form'>
					<input aria-label='Host title' value={title} onChange={(e) => setTitle(e.target.value)} />
					<button
						type='button'
						onClick={() =>
							report(
								'properties',
								actions.documentProperties.apply({
									core: { title },
									custom: { set: [{ name: 'Reviewer', value: 'Host', type: 'lpwstr' }] },
								}),
							)
						}
					>
						Apply properties
					</button>
					<button
						type='button'
						onClick={() =>
							report(
								'properties',
								actions.documentProperties.apply({
									custom: { set: [{ name: 'Count', value: 'many', type: 'i4' }] },
								}),
							)
						}
					>
						Apply bad property
					</button>
				</div>
			)}
			{pending === 'fontEmbedding' && (
				<div role='group' aria-label='Host fonts form'>
					<button
						type='button'
						onClick={() => report('fonts', actions.fontEmbedding.setEmbedFonts(true))}
					>
						Embed fonts
					</button>
					<button
						type='button'
						onClick={() => report('fonts', actions.fontEmbedding.setEmbedFonts(false))}
					>
						Do not embed fonts
					</button>
				</div>
			)}
			{pending === 'digitalSignatures' && (
				<div role='group' aria-label='Host signatures form'>
					<button
						type='button'
						onClick={() => report('signatures', actions.digitalSignatures.remove())}
					>
						Remove signatures
					</button>
					<button
						type='button'
						onClick={() => report('signatures', actions.digitalSignatures.sign())}
					>
						Sign
					</button>
				</div>
			)}
		</aside>
	);
}
