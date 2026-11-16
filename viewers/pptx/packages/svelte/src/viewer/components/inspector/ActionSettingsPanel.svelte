<script lang="ts">
	/**
	 * ActionSettingsPanel: PowerPoint's Insert > Action dialog as an inspector
	 * card. Each trigger (`actionClick` / `actionHover`) is the shared
	 * `ACTION_SETTINGS_SECTIONS` drawn by `InspectorFields` over its context
	 * (`actionSettingsTargetOf`: its action, the pending pick and the deck's
	 * slides and custom shows); a change holds its type as pending and, once it
	 * can carry its target, commits through core's `elementActionToPptxAction`.
	 */
	import type { ElementActionType, PptxElement } from 'pptx-viewer-core';
	import { elementActionToPptxAction } from 'pptx-viewer-core';
	import { ACTION_SETTINGS_SECTIONS, actionSettingsTargetOf } from 'ooxml-ui/pptx';
	import type { ActionSettingsChange } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import InspectorFields from './InspectorFields.svelte';

	const { editor, el }: { editor: EditorState; el: PptxElement } = $props();
	const t = useTranslator();
	const triggers = ['click', 'hover'] as const;

	/** The type picked but not committed, per trigger. */
	let pending = $state<Partial<Record<'click' | 'hover', ElementActionType>>>({});

	function targetOf(trigger: 'click' | 'hover') {
		return actionSettingsTargetOf(
			el,
			trigger,
			pending[trigger],
			editor.slides,
			editor.customShows.map((show) => ({ id: show.id, name: show.name })),
		);
	}

	function apply(trigger: 'click' | 'hover', change: ActionSettingsChange): void {
		pending = { ...pending, [trigger]: change.pendingType };
		if (change.commit) {
			const action = elementActionToPptxAction(change.commit);
			editor.applyElementPatch(
				el.id,
				(trigger === 'click' ? { actionClick: action } : { actionHover: action }) as Partial<PptxElement>,
			);
		}
	}
</script>

<!-- Keyed on the element so a half-made pick does not follow the inspector to
     the next shape, which never had it. -->
{#key el.id}
	<div class="pptx-svelte-action-settings" data-pptx-action-settings>
		{#each triggers as trigger (trigger)}
			<div class="pptx-svelte-action-trigger" data-pptx-action-trigger={trigger}>
				<span>{trigger === 'click' ? t('pptx.action.onClick') : t('pptx.action.onHover')}</span>
				<InspectorFields
					sections={ACTION_SETTINGS_SECTIONS}
					target={targetOf(trigger)}
					disabled={!editor.editable}
					onpatch={(change) => apply(trigger, change)}
				/>
			</div>
		{/each}
	</div>
{/key}

<style>
	.pptx-svelte-action-settings {
		display: grid;
		gap: 10px;
	}
	.pptx-svelte-action-trigger > span {
		font-weight: 600;
	}
</style>
