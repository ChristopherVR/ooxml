<script lang="ts">
	import type {
		SmartArtColorScheme,
		SmartArtLayoutType,
		SmartArtStyle,
		SmartArtPptxElement,
	} from 'pptx-viewer-core';
	import { SWITCHABLE_LAYOUT_TYPES } from 'pptx-viewer-core';
	import {
		reflowSmartArtData,
		schemaLabel,
		SMARTART_COLOR_SCHEME_LABEL_KEYS,
		SMARTART_STYLE_LABEL_KEYS,
		switchSmartArtLayoutData,
	} from 'ooxml-ui/pptx';
	import type { SmartArtTextPaneChangeEvent, SmartArtTextPaneViewState } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';

	const { editor, el }: { editor: EditorState; el: SmartArtPptxElement } = $props();
	const t = useTranslator();
	const data = $derived(el.smartArtData);

	// `dgm:colorsDef` / `dgm:styleDef` family tokens. Both lists stay explicit so
	// spelling them out through the shared tables cannot change which variations
	// the editor offers.
	const colorSchemes: readonly SmartArtColorScheme[] = [
		'colorful1',
		'colorful2',
		'colorful3',
		'monochromatic1',
		'monochromatic2',
	];
	const diagramStyles: readonly SmartArtStyle[] = ['flat', 'moderate', 'intense'];

	// The text pane is the shared `pptx-ui-smartart-text-pane` element (outline with levels,
	// keyboard structure edits, move, sub-item, remove, per-node style); this section feeds it
	// the data and commits what it emits.
	const paneState = $derived<SmartArtTextPaneViewState>({
		smartArtData: data,
		canEdit: editor.editable,
		translate: t,
	});

	function setLayout(layout: SmartArtLayoutType): void {
		if (data && layout !== data.resolvedLayoutType) {
			applyData(switchSmartArtLayoutData(data, layout));
		}
	}

	function setColorScheme(scheme: SmartArtColorScheme): void {
		if (data) {
			applyData({ ...data, colorScheme: scheme });
		}
	}
	/**
	 * Single commit funnel, mirroring React's `applySmartArtData`. Every edit
	 * here can clear the cached `dsp` drawing (add / remove / promote / demote /
	 * reorder / style / layout switch all do), so each one is reflowed: without
	 * it the diagram fell back to the crude family approximation. The reflow is
	 * a no-op while the cached drawing survives, so it never overrides it.
	 */
	function applyData(next: NonNullable<typeof data>): void {
		const box = { width: el.width, height: el.height };
		editor.applyElementPatch(el.id, { smartArtData: reflowSmartArtData(next, el.id, box) });
	}
	function setDiagramStyle(style: SmartArtStyle): void {
		if (data) {
			applyData({ ...data, style, drawingDirty: true, drawingShapes: [] });
		}
	}
	function onPaneChange(event: Event): void {
		const { smartArtData: next, recentColor } = (event as SmartArtTextPaneChangeEvent).detail;
		applyData(next);
		if (recentColor) {
			editor.recordRecentColor(recentColor);
		}
	}
</script>

{#if data}
	<span class="pptx-svelte-smartart-label">{t('pptx.smartart.switchLayout')}</span>
	<div class="pptx-svelte-smartart-layouts">
		{#each SWITCHABLE_LAYOUT_TYPES as layout}
			<button
				type="button"
				class:active={data.resolvedLayoutType === layout}
				aria-pressed={data.resolvedLayoutType === layout}
				data-testid={`smartart-layout-${layout}`}
				onclick={() => setLayout(layout)}
			>
				{t(`pptx.smartart.category.${layout}`)}
			</button>
		{/each}
	</div>

	<label class="pptx-svelte-smartart-field">
		<span>{t('pptx.smartart.colorScheme')}</span>
		<pptx-ui-select
			aria-label={t('pptx.smartart.colorScheme')}
			data-testid="smartart-color-scheme"
			value={data.colorScheme ?? 'colorful1'}
			onchange={(event) => setColorScheme(event.currentTarget.value as SmartArtColorScheme)}
		>
			{#each colorSchemes as scheme}
				<option value={scheme}>{schemaLabel(SMARTART_COLOR_SCHEME_LABEL_KEYS, scheme, t)}</option>
			{/each}
		</pptx-ui-select>
	</label>
	<label class="pptx-svelte-smartart-field"><span>{t('pptx.smartart.style')}</span><pptx-ui-select aria-label={t('pptx.smartart.style')} value={data.style ?? 'moderate'} onchange={(event) => setDiagramStyle(event.currentTarget.value as SmartArtStyle)}>{#each diagramStyles as diagramStyle}<option value={diagramStyle}>{schemaLabel(SMARTART_STYLE_LABEL_KEYS, diagramStyle, t)}</option>{/each}</pptx-ui-select></label>

	<pptx-ui-smartart-text-pane
		data-testid="smartart-text-pane"
		state={paneState}
		onpptx-smartart-data-change={onPaneChange}
	></pptx-ui-smartart-text-pane>
{/if}

<style>
	.pptx-svelte-smartart-label {
		display: block;
		margin-bottom: 6px;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-smartart-layouts {
		display: grid;
		grid-template-columns: repeat(3, minmax(0, 1fr));
		gap: 5px;
		margin-bottom: 10px;
	}

	.pptx-svelte-smartart-layouts button {
		min-width: 0;
		padding: 6px 3px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: transparent;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		font: inherit;
		font-size: 10px;
		cursor: pointer;
		overflow: hidden;
		text-overflow: ellipsis;
	}

	.pptx-svelte-smartart-layouts button:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
		color: var(--pptx-accent-foreground, var(--office-foreground, #1f2937));
	}

	.pptx-svelte-smartart-layouts button.active {
		border-color: var(--pptx-primary, var(--office-accent, #2563eb));
		background: color-mix(in srgb, var(--pptx-primary, var(--office-accent, #2563eb)) 16%, transparent);
		color: var(--pptx-primary, var(--office-accent, #2563eb));
	}

	.pptx-svelte-smartart-field {
		display: flex;
		align-items: center;
		justify-content: space-between;
		gap: 8px;
		margin-bottom: 10px;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.pptx-svelte-smartart-field select {
		height: 26px;
		box-sizing: border-box;
		padding: 2px 6px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-background, var(--office-background, #ffffff));
		color: var(--pptx-foreground, var(--office-foreground, #1f2937));
		font: inherit;
	}
</style>
