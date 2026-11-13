import type { PptxCustomProperty } from 'ooxml-core/pptx';

import {
	BUILDING_BLOCK_OK,
	buildingBlockFailure,
	type BuildingBlockResult,
} from './building-block-actions';
import { isMarkedAsFinal, type ReadOnlyRecommendation } from './read-only-recommendation';

/**
 * Framework-free password-protection workflow of a host-owned shell: what the deck is protected
 * with, and the operations that change it. Stored passwords are never part of the state.
 */

/** Open passwords shorter than this are refused (the bundled dialog's rule). */
export const MIN_OPEN_PASSWORD_LENGTH = 4;

/** `docProps/custom.xml`'s well-known "Mark as Final" property. */
const MARK_AS_FINAL_PROPERTY = '_MarkAsFinal';

/** A strength score from 0 (empty or trivial) to 4, for a strength meter. */
export function passwordStrength(password: string): number {
	if (!password) {
		return 0;
	}
	let score = 0;
	if (password.length >= 8) {
		score++;
	}
	if (password.length >= 12) {
		score++;
	}
	if (/[A-Z]/u.test(password) && /[a-z]/u.test(password)) {
		score++;
	}
	if (/\d/u.test(password)) {
		score++;
	}
	if (/[^A-Za-z0-9]/u.test(password)) {
		score++;
	}
	return Math.min(score, 4);
}

/** The bundled dialog's checks, in its order: required, confirmed, long enough. */
export function validateOpenPassword(password: string, confirmation: string): BuildingBlockResult {
	if (!password) {
		return buildingBlockFailure('password-required');
	}
	if (password !== confirmation) {
		return buildingBlockFailure('password-mismatch');
	}
	return password.length < MIN_OPEN_PASSWORD_LENGTH
		? buildingBlockFailure('password-too-short')
		: BUILDING_BLOCK_OK;
}

/** What the deck is protected with, as state only (never a password or hash). */
export interface PasswordProtectionState {
	/** Saving encrypts the file with an open password set in this session. */
	openPassword: boolean;
	/** The loaded deck carries a modify password (`p:modifyVerifier` with a hash). */
	modifyPassword: boolean;
	/** The loaded deck asks to be opened read-only (a modify verifier or Mark as Final). */
	readOnlyRecommended: boolean;
	/** The deck is marked as final (`_MarkAsFinal`). */
	markedFinal: boolean;
	/**
	 * Which protections the operations can change. A modify password and the read-only
	 * recommendation are only reported: the viewer has no way to write them, it preserves what the
	 * deck arrived with.
	 */
	supported: {
		openPassword: true;
		markedFinal: true;
		modifyPassword: false;
		readOnlyRecommended: false;
	};
}

export interface PasswordProtectionInput {
	openPasswordSet: boolean;
	recommendation: Pick<ReadOnlyRecommendation, 'kind'>;
	hasModifyVerifier: boolean;
	customProperties: readonly PptxCustomProperty[] | undefined;
}

export function describePasswordProtection(
	input: PasswordProtectionInput,
): PasswordProtectionState {
	return {
		openPassword: input.openPasswordSet,
		modifyPassword: input.hasModifyVerifier,
		readOnlyRecommended: input.recommendation.kind !== null,
		markedFinal: isMarkedAsFinal(input.customProperties),
		supported: {
			openPassword: true,
			markedFinal: true,
			modifyPassword: false,
			readOnlyRecommended: false,
		},
	};
}

/** The custom properties with Mark as Final switched on or off (`undefined` when nothing changes). */
export function markAsFinalProperties(
	customProperties: readonly PptxCustomProperty[],
	final: boolean,
): PptxCustomProperty[] | undefined {
	if (isMarkedAsFinal(customProperties) === final) {
		return undefined;
	}
	const others = customProperties.filter((property) => property.name !== MARK_AS_FINAL_PROPERTY);
	return final
		? [...others, { name: MARK_AS_FINAL_PROPERTY, value: 'true', type: 'bool' }]
		: others;
}

export interface PasswordProtectionActions {
	/** The current protections (state only). */
	state: PasswordProtectionState;
	/** Encrypt on save with this open password (validated like the bundled dialog). */
	setOpenPassword: (password: string, confirmation: string) => BuildingBlockResult;
	/** Save without an open password. */
	clearOpenPassword: () => BuildingBlockResult;
	/** Mark or unmark the deck as final. */
	setMarkedFinal: (final: boolean) => BuildingBlockResult;
	/**
	 * Not supported: the viewer cannot write a modify password or the read-only recommendation.
	 * Always refused with `unsupported` (or the permission code when the deck is not editable).
	 */
	setModifyPassword: (password: string | null) => BuildingBlockResult;
}

export interface PasswordProtectionActionsInput {
	state: PasswordProtectionState;
	customProperties: () => readonly PptxCustomProperty[];
	/** The permission check, run again at every operation. */
	refuse: () => BuildingBlockResult | undefined;
	applyOpenPassword: (password: string | null) => void;
	applyCustomProperties: (next: PptxCustomProperty[]) => void;
	/** The deck changed (the saved file will differ). */
	markDirty: () => void;
	/** The workflow finished successfully (closes the pending request). */
	complete: (result: BuildingBlockResult) => BuildingBlockResult;
}

export function buildPasswordProtectionActions(
	input: PasswordProtectionActionsInput,
): PasswordProtectionActions {
	const guard = (run: () => BuildingBlockResult): BuildingBlockResult => input.refuse() ?? run();
	return {
		state: input.state,
		setOpenPassword: (password, confirmation) =>
			guard(() => {
				const checked = validateOpenPassword(password, confirmation);
				if (!checked.ok) {
					return checked;
				}
				input.applyOpenPassword(password);
				input.markDirty();
				return input.complete(BUILDING_BLOCK_OK);
			}),
		clearOpenPassword: () =>
			guard(() => {
				input.applyOpenPassword(null);
				input.markDirty();
				return input.complete(BUILDING_BLOCK_OK);
			}),
		setMarkedFinal: (final) =>
			guard(() => {
				const next = markAsFinalProperties(input.customProperties(), final);
				if (next) {
					input.applyCustomProperties(next);
				}
				return input.complete(BUILDING_BLOCK_OK);
			}),
		setModifyPassword: () => input.refuse() ?? buildingBlockFailure('unsupported'),
	};
}
