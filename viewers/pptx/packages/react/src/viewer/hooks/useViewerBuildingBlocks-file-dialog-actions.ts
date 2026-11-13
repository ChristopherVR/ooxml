import type {
	BuildingBlockGroupPorts,
	BuildingBlockRequest,
	ReadOnlyRecommendation,
} from 'ooxml-ui/pptx';
import { useRef } from 'react';

import { useDialogFlagRequest } from './useDialogFlagRequest';
import type { EditorHistoryResult } from './useEditorHistory';
import type { PropertyHandlersResult } from './usePropertyHandlers';
import type { ViewerDialogsResult } from './useViewerDialogs';
import type { ViewerState } from './useViewerState';

export type FileDialogPorts = Pick<
	BuildingBlockGroupPorts,
	'passwordProtection' | 'documentProperties' | 'fontEmbedding' | 'digitalSignatures'
>;

export interface UseFileDialogPortsInput {
	state: ViewerState;
	dialogs: ViewerDialogsResult;
	propertyHandlers: PropertyHandlersResult;
	history: EditorHistoryResult;
	content: ArrayBuffer | Uint8Array | null | undefined;
	recommendation: Pick<ReadOnlyRecommendation, 'kind'>;
	hasModifyVerifier: boolean;
	/** Raise a host request through the shared session (viewable kinds fire read-only too). */
	request: (next: BuildingBlockRequest) => void;
}

/**
 * Password protection, document properties, font embedding and digital signatures: the ports of
 * the shared action groups (`ooxml-ui/pptx`, `render/building-blocks/groups`) over this binding's
 * state, plus the effects that turn the toolbar's dialog flags into host requests. The ports are
 * created once and read the latest render through a ref.
 */
export function useFileDialogPorts(input: UseFileDialogPortsInput): FileDialogPorts {
	const latest = useRef(input);
	latest.current = input;
	const { dialogs, request } = input;
	useDialogFlagRequest(dialogs.isPasswordDialogOpen, () => {
		dialogs.setIsPasswordDialogOpen(false);
		request({ kind: 'passwordProtection' });
	});
	useDialogFlagRequest(dialogs.isDocPropsDialogOpen, () => {
		dialogs.setIsDocPropsDialogOpen(false);
		request({ kind: 'documentProperties' });
	});
	useDialogFlagRequest(dialogs.isFontEmbeddingOpen, () => {
		dialogs.setIsFontEmbeddingOpen(false);
		request({ kind: 'fontEmbedding' });
	});
	useDialogFlagRequest(dialogs.isDigitalSigDialogOpen, () => {
		dialogs.setIsDigitalSigDialogOpen(false);
		request({ kind: 'digitalSignatures' });
	});

	// Created once; every member reads the latest render through `latest`.
	const ports = useRef<FileDialogPorts | null>(null);
	ports.current ??= (() => {
		const now = () => latest.current;
		const markDirty = () => now().history.markDirty();
		return {
			passwordProtection: {
				input: () => ({
					openPasswordSet: now().dialogs.isPasswordProtected,
					recommendation: now().recommendation,
					hasModifyVerifier: now().hasModifyVerifier,
					customProperties: now().state.customProperties,
				}),
				customProperties: () => now().state.customProperties,
				applyOpenPassword: (password) =>
					password === null
						? now().dialogs.handleRemovePassword()
						: now().dialogs.handleSetPassword(password),
				applyCustomProperties: (next) => now().propertyHandlers.handleUpdateCustomProperties(next),
				markDirty,
			},
			documentProperties: {
				current: () => ({
					core: now().state.coreProperties ?? {},
					custom: now().state.customProperties,
					app: now().state.appProperties ?? {},
				}),
				applyUpdate: (update) => {
					const handlers = now().propertyHandlers;
					if (update.core) {
						handlers.handleUpdateCoreProperties(update.core);
					}
					if (update.custom) {
						handlers.handleUpdateCustomProperties(update.custom);
					}
					if (update.app) {
						handlers.handleUpdateAppProperties(update.app);
					}
				},
			},
			fontEmbedding: {
				input: () => ({
					embedFonts: now().dialogs.embedFontsEnabled,
					embeddedFamilies: now().state.embeddedFonts.map((font) => font.name),
					usedFamilies: now().propertyHandlers.usedFontFamilies,
					descriptor: now().dialogs.fontEmbedding,
				}),
				apply: (enabled) => now().dialogs.setEmbedFontsEnabled(enabled),
				markDirty,
			},
			digitalSignatures: {
				signed: () => now().state.hasDigitalSignatures,
				count: () => now().state.digitalSignatureCount,
				content: () => now().content,
				markDirty,
			},
		};
	})();
	return ports.current;
}
