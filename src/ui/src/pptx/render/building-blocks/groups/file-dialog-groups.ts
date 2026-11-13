/**
 * The file-level dialogs the building blocks never mount (Password Protection, Document
 * Properties, Font Embedding, Digital Signatures) as action groups. The decisions are the shared
 * `building-block-*` modules; each group only binds them to the session's gating and request
 * completion, and to a small port the binding implements over its own state.
 *
 * @module render/building-blocks/groups/file-dialog-groups
 */
import type { PptxCustomProperty } from 'ooxml-core/pptx';

import {
	buildDocumentPropertiesActions,
	type DocumentPropertiesSnapshot,
	type DocumentPropertiesUpdate,
} from '../../building-block-document-properties';
import {
	buildFontEmbeddingActions,
	describeFontEmbeddingState,
} from '../../building-block-font-embedding';
import {
	buildPasswordProtectionActions,
	describePasswordProtection,
	type PasswordProtectionInput,
} from '../../building-block-password';
import { buildSignatureActions, describeSignaturesState } from '../../building-block-signatures';
import { defineBuildingBlockActionGroup } from '../action-groups';

export interface PasswordProtectionPort {
	/** The protections as the binding knows them now. */
	input(): PasswordProtectionInput;
	customProperties(): readonly PptxCustomProperty[];
	/** Set (a password) or clear (null) the open password the next save encrypts with. */
	applyOpenPassword(password: string | null): void;
	applyCustomProperties(next: PptxCustomProperty[]): void;
	markDirty(): void;
}

export const passwordProtectionGroup = defineBuildingBlockActionGroup({
	id: 'passwordProtection',
	create: (context, port: PasswordProtectionPort) =>
		buildPasswordProtectionActions({
			state: describePasswordProtection(port.input()),
			customProperties: () => port.customProperties(),
			refuse: () => context.refuse('passwordProtection'),
			applyOpenPassword: (password) => port.applyOpenPassword(password),
			applyCustomProperties: (next) => port.applyCustomProperties(next),
			markDirty: () => port.markDirty(),
			complete: (result) => context.complete('passwordProtection', result),
		}),
});

export interface DocumentPropertiesPort {
	current(): DocumentPropertiesSnapshot;
	/** Store the changed parts (history and dirty as the bundled dialog's OK). */
	applyUpdate(update: DocumentPropertiesUpdate): void;
}

export const documentPropertiesGroup = defineBuildingBlockActionGroup({
	id: 'documentProperties',
	create: (context, port: DocumentPropertiesPort) =>
		buildDocumentPropertiesActions({
			current: () => port.current(),
			refuse: () => context.refuse('documentProperties'),
			applyUpdate: (update) => port.applyUpdate(update),
			complete: (result) => context.complete('documentProperties', result),
		}),
});

export interface FontEmbeddingPort {
	/** What `describeFontEmbeddingState` reads, now. */
	input(): Parameters<typeof describeFontEmbeddingState>[0];
	/** Keep (true) or strip (false) the embedded fonts on save. */
	apply(enabled: boolean): void;
	markDirty(): void;
}

export const fontEmbeddingGroup = defineBuildingBlockActionGroup({
	id: 'fontEmbedding',
	create: (context, port: FontEmbeddingPort) =>
		buildFontEmbeddingActions({
			state: describeFontEmbeddingState(port.input()),
			refuse: () => context.refuse('fontEmbedding'),
			latest: () => {
				const { embedFonts, canEmbed } = describeFontEmbeddingState(port.input());
				return { embedFonts, canEmbed };
			},
			apply: (enabled) => port.apply(enabled),
			markDirty: () => port.markDirty(),
			complete: (result) => context.complete('fontEmbedding', result),
		}),
});

export interface DigitalSignaturesPort {
	signed(): boolean;
	count(): number;
	/** The loaded package's bytes, which the signatures are read from. */
	content(): ArrayBuffer | Uint8Array | null | undefined;
	markDirty(): void;
}

export const digitalSignaturesGroup = defineBuildingBlockActionGroup({
	id: 'digitalSignatures',
	create: (context, port: DigitalSignaturesPort) =>
		buildSignatureActions({
			state: describeSignaturesState({ signed: port.signed(), count: port.count() }),
			content: () => port.content(),
			refuse: () => context.refuse('digitalSignatures'),
			markDirty: () => port.markDirty(),
			complete: (result) => context.complete('digitalSignatures', result),
		}),
});
