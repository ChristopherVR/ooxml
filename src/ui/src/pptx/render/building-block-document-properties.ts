import type { PptxAppProperties, PptxCoreProperties, PptxCustomProperty } from 'ooxml-core/pptx';

import {
	BUILDING_BLOCK_OK,
	buildingBlockFailure,
	type BuildingBlockResult,
} from './building-block-actions';

/**
 * Framework-free document-properties workflow: the Summary fields, the custom properties and the
 * editable Statistics fields of the bundled Document Properties dialog, as a patchable snapshot.
 * The bundled dialog uses {@link documentPropertiesChanged} for its dirty check, so both agree on
 * what counts as a change.
 */

/** The core properties the dialog's Summary tab edits. */
export const EDITABLE_CORE_PROPERTY_KEYS = [
	'title',
	'subject',
	'creator',
	'keywords',
	'description',
	'category',
] as const satisfies ReadonlyArray<keyof PptxCoreProperties>;
export type EditableCorePropertyKey = (typeof EDITABLE_CORE_PROPERTY_KEYS)[number];

/** The application properties the Statistics tab lets the user edit. */
export const EDITABLE_APP_PROPERTY_KEYS = ['company', 'manager'] as const;
export type EditableAppPropertyKey = (typeof EDITABLE_APP_PROPERTY_KEYS)[number];

/** The custom property types the Custom tab offers. */
export const CUSTOM_PROPERTY_TYPES = ['lpwstr', 'i4', 'filetime', 'bool'] as const;

export interface DocumentPropertiesSnapshot {
	core: PptxCoreProperties;
	custom: PptxCustomProperty[];
	/** The whole application part (statistics are read-only; only company and manager patch). */
	app: PptxAppProperties;
}

/** A partial change. Absent groups and fields stay as they are; an empty string clears a field. */
export interface DocumentPropertiesPatch {
	core?: Partial<Record<EditableCorePropertyKey, string>>;
	app?: Partial<Record<EditableAppPropertyKey, string>>;
	custom?: {
		/** Add, or replace by name. */
		set?: PptxCustomProperty[];
		/** Remove by name. */
		remove?: string[];
	};
}

/** What changed, shaped as the viewer's three property setters take it. */
export interface DocumentPropertiesUpdate {
	core?: Partial<PptxCoreProperties>;
	custom?: PptxCustomProperty[];
	app?: Partial<PptxAppProperties>;
}

/** Whether `draft` differs from `current` in a field the dialog edits. */
export function documentPropertiesChanged(
	current: {
		core: PptxCoreProperties;
		custom: readonly PptxCustomProperty[];
		app?: PptxAppProperties;
	},
	draft: {
		core: PptxCoreProperties;
		custom: readonly PptxCustomProperty[];
		app?: Partial<PptxAppProperties>;
	},
): boolean {
	return (
		EDITABLE_CORE_PROPERTY_KEYS.some(
			(key) => (draft.core[key] ?? '') !== (current.core[key] ?? ''),
		) ||
		appPropertiesChanged(current.app, draft.app) ||
		draft.custom.length !== current.custom.length ||
		draft.custom.some((property, index) => {
			const existing = current.custom[index];
			return (
				property.name !== existing?.name ||
				property.value !== existing?.value ||
				property.type !== existing?.type
			);
		})
	);
}

/** Whether the editable Statistics fields (company, manager) differ. */
export function appPropertiesChanged(
	current: PptxAppProperties | undefined,
	draft: Partial<PptxAppProperties> | undefined,
): boolean {
	return EDITABLE_APP_PROPERTY_KEYS.some((key) => (draft?.[key] ?? '') !== (current?.[key] ?? ''));
}

function validCustomProperty(property: PptxCustomProperty): boolean {
	if (typeof property.name !== 'string' || property.name.trim() === '') {
		return false;
	}
	if (!(CUSTOM_PROPERTY_TYPES as readonly string[]).includes(property.type)) {
		return false;
	}
	if (typeof property.value !== 'string') {
		return false;
	}
	switch (property.type) {
		case 'i4':
			return /^-?\d+$/u.test(property.value.trim());
		case 'bool':
			return property.value === 'true' || property.value === 'false';
		case 'filetime':
			return !Number.isNaN(Date.parse(property.value));
		default:
			return true;
	}
}

export type DocumentPropertiesPlan =
	| ({ ok: true; changed: boolean } & DocumentPropertiesUpdate)
	| { ok: false; code: 'invalid-data' };

/**
 * Apply a patch to a snapshot without touching it: the updates to hand the viewer's setters, or
 * `invalid-data` when a field is not a string or a custom property is malformed (blank name, an
 * unknown type, a value that does not parse as its type).
 */
export function planDocumentPropertiesPatch(
	current: DocumentPropertiesSnapshot,
	patch: DocumentPropertiesPatch,
): DocumentPropertiesPlan {
	const invalid = { ok: false, code: 'invalid-data' } as const;
	const coreEntries = Object.entries(patch.core ?? {});
	const appEntries = Object.entries(patch.app ?? {});
	for (const [key, value] of coreEntries) {
		if (
			!(EDITABLE_CORE_PROPERTY_KEYS as readonly string[]).includes(key) ||
			typeof value !== 'string'
		) {
			return invalid;
		}
	}
	for (const [key, value] of appEntries) {
		if (
			!(EDITABLE_APP_PROPERTY_KEYS as readonly string[]).includes(key) ||
			typeof value !== 'string'
		) {
			return invalid;
		}
	}
	const set = patch.custom?.set ?? [];
	if (!set.every(validCustomProperty)) {
		return invalid;
	}
	const core = Object.fromEntries(coreEntries) as Partial<PptxCoreProperties>;
	const app = Object.fromEntries(appEntries) as Partial<PptxAppProperties>;
	const hasCustom = set.length > 0 || (patch.custom?.remove?.length ?? 0) > 0;
	let custom = current.custom;
	if (hasCustom) {
		const removed = new Set(patch.custom?.remove ?? []);
		// Replaced properties keep their place; new ones are appended.
		custom = [
			...current.custom.flatMap((property) => {
				if (removed.has(property.name)) {
					return [];
				}
				const replacement = set.find((next) => next.name === property.name);
				return [replacement ? { ...replacement } : property];
			}),
			...set.filter((next) => !current.custom.some((property) => property.name === next.name)),
		];
	}
	const changed = documentPropertiesChanged(
		{ core: current.core, custom: current.custom, app: current.app },
		{ core: { ...current.core, ...core }, custom, app: { ...current.app, ...app } },
	);
	return {
		ok: true,
		changed,
		...(coreEntries.length > 0 ? { core } : {}),
		...(hasCustom ? { custom } : {}),
		...(appPropertiesChanged(current.app, { ...current.app, ...app }) ? { app } : {}),
	};
}

export interface DocumentPropertiesActions {
	/** The current properties (copies; editing them changes nothing). */
	current: DocumentPropertiesSnapshot;
	/** Apply a patch. Validated; nothing is applied when any part is invalid. */
	apply: (patch: DocumentPropertiesPatch) => BuildingBlockResult;
}

export interface DocumentPropertiesActionsInput {
	current: () => DocumentPropertiesSnapshot;
	refuse: () => BuildingBlockResult | undefined;
	applyUpdate: (update: DocumentPropertiesUpdate) => void;
	complete: (result: BuildingBlockResult) => BuildingBlockResult;
}

function copy(snapshot: DocumentPropertiesSnapshot): DocumentPropertiesSnapshot {
	return {
		core: { ...snapshot.core },
		custom: snapshot.custom.map((property) => ({ ...property })),
		app: { ...snapshot.app },
	};
}

export function buildDocumentPropertiesActions(
	input: DocumentPropertiesActionsInput,
): DocumentPropertiesActions {
	return {
		current: copy(input.current()),
		apply: (patch) => {
			const denied = input.refuse();
			if (denied) {
				return denied;
			}
			const plan = planDocumentPropertiesPatch(input.current(), patch);
			if (!plan.ok) {
				return buildingBlockFailure(plan.code);
			}
			if (plan.changed) {
				input.applyUpdate({
					...(plan.core ? { core: plan.core } : {}),
					...(plan.custom ? { custom: plan.custom } : {}),
					...(plan.app ? { app: plan.app } : {}),
				});
			}
			return input.complete(BUILDING_BLOCK_OK);
		},
	};
}
