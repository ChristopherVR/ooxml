import {
	buildDocumentPropertiesActions,
	buildFontEmbeddingActions,
	buildPasswordProtectionActions,
	buildSignatureActions,
	describeFontEmbeddingState,
	describePasswordProtection,
	describeSignaturesState,
} from 'ooxml-ui/pptx';
import type {
	BuildingBlockActionId,
	BuildingBlockResult,
	DocumentPropertiesActions,
	FontEmbeddingActions,
	PasswordProtectionActions,
	ReadOnlyRecommendation,
	SignatureActions,
} from 'ooxml-ui/pptx';
import { useEffect, useRef } from 'react';

import type { EditorHistoryResult } from './useEditorHistory';
import type { PropertyHandlersResult } from './usePropertyHandlers';
import type { BuildingBlockRequest } from './useViewerBuildingBlocks-actions-types';
import type { ViewerDialogsResult } from './useViewerDialogs';
import type { ViewerState } from './useViewerState';

/** The four file-level dialogs the building blocks never mount, as host operations. */
export interface FileDialogActions {
	passwordProtection: PasswordProtectionActions;
	documentProperties: DocumentPropertiesActions;
	fontEmbedding: FontEmbeddingActions;
	digitalSignatures: SignatureActions;
}

export interface UseFileDialogActionsInput {
	state: ViewerState;
	dialogs: ViewerDialogsResult;
	propertyHandlers: PropertyHandlersResult;
	history: EditorHistoryResult;
	content: ArrayBuffer | Uint8Array | null | undefined;
	recommendation: Pick<ReadOnlyRecommendation, 'kind'>;
	hasModifyVerifier: boolean;
	/** The shared permission check of an action (null when it may run). */
	refuse: (id: BuildingBlockActionId) => BuildingBlockResult | undefined;
	/** Raise a request; `viewable` lets it fire while the deck is read-only or protected. */
	request: (next: BuildingBlockRequest, id: BuildingBlockActionId, viewable?: boolean) => void;
	complete: (
		kind: BuildingBlockRequest['kind'],
		result: BuildingBlockResult,
	) => BuildingBlockResult;
}

/**
 * Password protection, document properties, font embedding and digital signatures for a host's
 * own UI. The toolbar controls that used to open internal dialogs now raise a request; the host
 * reads the state and applies through these groups. The decisions are the shared
 * `building-block-*` modules in `ooxml-ui/pptx`; this hook only maps them onto the viewer's state.
 */
/**
 * A toolbar control opened an internal dialog by flag: clear the flag and raise the host request.
 * Viewing is allowed read-only; applying is refused by the operations themselves.
 */
function useOpenFlagRequest(isOpen: boolean, setOpen: (open: boolean) => void, raise: () => void) {
	const raiseRef = useRef(raise);
	raiseRef.current = raise;
	useEffect(() => {
		if (isOpen) {
			setOpen(false);
			raiseRef.current();
		}
	}, [isOpen, setOpen]);
}

export function useFileDialogActions(input: UseFileDialogActionsInput): FileDialogActions {
	const { state, dialogs, propertyHandlers } = input;
	// Operations run long after the render that made them: they read the latest of everything.
	const latest = useRef(input);
	latest.current = input;
	const refuse = (id: BuildingBlockActionId) => latest.current.refuse(id);

	// The toolbar opens the internal dialogs by flag; nothing renders them here.
	const { request } = input;
	useOpenFlagRequest(dialogs.isPasswordDialogOpen, dialogs.setIsPasswordDialogOpen, () =>
		request({ kind: 'passwordProtection' }, 'passwordProtection', true),
	);
	useOpenFlagRequest(dialogs.isDocPropsDialogOpen, dialogs.setIsDocPropsDialogOpen, () =>
		request({ kind: 'documentProperties' }, 'documentProperties', true),
	);
	useOpenFlagRequest(dialogs.isFontEmbeddingOpen, dialogs.setIsFontEmbeddingOpen, () =>
		request({ kind: 'fontEmbedding' }, 'fontEmbedding', true),
	);
	useOpenFlagRequest(dialogs.isDigitalSigDialogOpen, dialogs.setIsDigitalSigDialogOpen, () =>
		request({ kind: 'digitalSignatures' }, 'digitalSignatures', true),
	);

	const complete = (kind: BuildingBlockRequest['kind']) => (result: BuildingBlockResult) =>
		latest.current.complete(kind, result);
	const markDirty = () => latest.current.history.markDirty();

	const embeddedFamilies = state.embeddedFonts.map((font) => font.name);
	return {
		passwordProtection: buildPasswordProtectionActions({
			state: describePasswordProtection({
				openPasswordSet: dialogs.isPasswordProtected,
				recommendation: input.recommendation,
				hasModifyVerifier: input.hasModifyVerifier,
				customProperties: state.customProperties,
			}),
			customProperties: () => latest.current.state.customProperties,
			refuse: () => refuse('passwordProtection'),
			applyOpenPassword: (password) =>
				password === null
					? latest.current.dialogs.handleRemovePassword()
					: latest.current.dialogs.handleSetPassword(password),
			applyCustomProperties: (next) =>
				latest.current.propertyHandlers.handleUpdateCustomProperties(next),
			markDirty,
			complete: complete('passwordProtection'),
		}),
		documentProperties: buildDocumentPropertiesActions({
			current: () => ({
				core: latest.current.state.coreProperties ?? {},
				custom: latest.current.state.customProperties,
				app: latest.current.state.appProperties ?? {},
			}),
			refuse: () => refuse('documentProperties'),
			applyUpdate: (update) => {
				const handlers = latest.current.propertyHandlers;
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
			complete: complete('documentProperties'),
		}),
		fontEmbedding: buildFontEmbeddingActions({
			state: describeFontEmbeddingState({
				embedFonts: dialogs.embedFontsEnabled,
				embeddedFamilies,
				usedFamilies: propertyHandlers.usedFontFamilies,
				descriptor: dialogs.fontEmbedding,
			}),
			refuse: () => refuse('fontEmbedding'),
			latest: () => ({
				embedFonts: latest.current.dialogs.embedFontsEnabled,
				canEmbed: latest.current.dialogs.fontEmbedding.interactive,
			}),
			apply: (enabled) => latest.current.dialogs.setEmbedFontsEnabled(enabled),
			markDirty,
			complete: complete('fontEmbedding'),
		}),
		digitalSignatures: buildSignatureActions({
			state: describeSignaturesState({
				signed: state.hasDigitalSignatures,
				count: state.digitalSignatureCount,
			}),
			content: () => latest.current.content,
			refuse: () => refuse('digitalSignatures'),
			markDirty,
			complete: complete('digitalSignatures'),
		}),
	};
}
