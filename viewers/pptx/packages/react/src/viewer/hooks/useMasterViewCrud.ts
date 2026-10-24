import type {
	CompatibilityWarningToast,
	MasterViewCrudAction,
	MasterViewCrudActionId,
	MasterViewTarget,
	SlideMasterHistoryStep,
} from 'ooxml-ui/pptx';
import {
	applyMasterViewCrudAction,
	historyStepUsesParts,
	masterViewCrudActions,
	masterViewCrudFailureKey,
	removedSlideMasterParts,
} from 'ooxml-ui/pptx';
import type { PptxData, PptxHandler, PptxSlide, PptxSlideMaster } from 'pptx-viewer-core';
import { useCallback, useMemo } from 'react';

import { useTranslation } from '../i18n/use-translation';
import { safePrompt } from '../utils/dom-helpers';

/**
 * useMasterViewCrud: Slide Master view sidebar CRUD (Insert/Duplicate/
 * Delete/Rename Layout and Slide Master).
 *
 * The shared `master-view-crud` module only ever reads `.slides` /
 * `.slideMasters` off the `PptxData` it is handed (verified against
 * `ooxml-core/src/pptx/core/builders/sdk/master-layout-crud*.ts`: every lookup
 * and `saveToZip` call touches only those two fields), so the minimal object
 * built here is safe even though React keeps the rest of a loaded deck's
 * fields split into their own state slots rather than one `PptxData`
 * (`useExportSaveAs`'s JSON export takes the same shortcut).
 */

export interface UseMasterViewCrudInput {
	handlerRef: React.RefObject<PptxHandler | null>;
	slides: PptxSlide[];
	slideMasters: PptxSlideMaster[];
	target: MasterViewTarget | null;
	setSlideMasters: React.Dispatch<React.SetStateAction<PptxSlideMaster[]>>;
	/** The Home > Layout list, which this binding keeps apart from the masters it comes from. */
	setLayoutOptions: (layoutOptions: NonNullable<PptxData['layoutOptions']>) => void;
	setActiveMasterIndex: React.Dispatch<React.SetStateAction<number>>;
	setActiveLayoutIndex: React.Dispatch<React.SetStateAction<number | null>>;
	markDirty: () => void;
	/** Drops the undo and redo steps that would bring back a part a command removed. */
	forgetHistory: (reaches: (step: SlideMasterHistoryStep) => boolean) => void;
	/** Surfaces a failure reason through the existing compat-toast stack. */
	pushToast: (toast: CompatibilityWarningToast) => void;
}

export interface UseMasterViewCrudResult {
	/** The sidebar's eight command buttons, resolved enabled/disabled for `target`. */
	crudActions: MasterViewCrudAction[];
	/** Run one command by id (rename* prompts for a name first). */
	handleCrudAction: (id: MasterViewCrudActionId) => Promise<void>;
}

export function useMasterViewCrud(input: UseMasterViewCrudInput): UseMasterViewCrudResult {
	const {
		handlerRef,
		slides,
		slideMasters,
		target,
		setSlideMasters,
		setLayoutOptions,
		setActiveMasterIndex,
		setActiveLayoutIndex,
		markDirty,
		forgetHistory,
		pushToast,
	} = input;
	const { t } = useTranslation();

	// See the module doc: only `.slides` / `.slideMasters` are ever read off
	// this by the shared helper or core's master-layout-crud SDK.
	const data = useMemo(
		() => ({ slides, slideMasters }) as unknown as PptxData,
		[slides, slideMasters],
	);

	const crudActions = useMemo(() => masterViewCrudActions(data, target), [data, target]);

	const handleCrudAction = useCallback(
		async (id: MasterViewCrudActionId) => {
			const handler = handlerRef.current;
			if (!handler || !target) {
				return;
			}
			const action = crudActions.find((candidate) => candidate.id === id);
			if (!action?.enabled) {
				return;
			}

			let name: string | undefined;
			if (id === 'renameLayout' || id === 'renameMaster') {
				const master = slideMasters[target.masterIndex];
				const current =
					id === 'renameLayout'
						? (master?.layouts?.[target.layoutIndex ?? -1]?.name ?? '')
						: (master?.name ?? '');
				const next = safePrompt(t('pptx.masterView.renamePrompt'), current)?.trim();
				if (!next) {
					return;
				}
				name = next;
			}

			const result = await applyMasterViewCrudAction(
				handler,
				data,
				id,
				target,
				name !== undefined ? { name } : undefined,
			);

			if (!result.ok) {
				pushToast({
					id: `master-crud-${Date.now()}`,
					code: `masterViewCrud:${result.reason}`,
					severity: 'warning',
					messageKey: masterViewCrudFailureKey(id, result.reason),
				});
				return;
			}

			// Adopt the helper's fresh handler and masters (it does not mutate the
			// ones it was given), and the layout list built from them, or Home >
			// Layout keeps offering a deleted layout and misses an inserted one. The slides stay as they are: the reloaded ones
			// carry no loaded pictures (each shows as a placeholder), and they hold
			// nothing the command changed. The previous handler stays alive: the
			// slides on screen and the undo stack still show its pictures.
			handlerRef.current = result.handler;
			setSlideMasters(result.data.slideMasters ?? []);
			setLayoutOptions(result.data.layoutOptions ?? []);
			setActiveMasterIndex(result.target.masterIndex);
			setActiveLayoutIndex(result.target.layoutIndex);
			const removed = removedSlideMasterParts(data, result.data);
			if (removed.length > 0) {
				forgetHistory((step) => historyStepUsesParts(step, removed));
			}
			markDirty();
		},
		[
			handlerRef,
			target,
			crudActions,
			slideMasters,
			t,
			data,
			pushToast,
			setSlideMasters,
			setLayoutOptions,
			setActiveMasterIndex,
			setActiveLayoutIndex,
			forgetHistory,
			markDirty,
		],
	);

	return { crudActions, handleCrudAction };
}
