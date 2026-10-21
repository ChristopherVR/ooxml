<script lang="ts">
	/**
	 * DocumentPropertiesSection: DOCUMENT card (Title / Author / Company /
	 * Application text fields + custom properties), the Svelte port of Vue's
	 * `DocumentPropertiesCard` (React `inspector/DocumentPropertiesCards.tsx`).
	 * Fields commit on change (blur / Enter) so history is not spammed per
	 * keystroke.
	 */
	import type { PptxAppProperties, PptxCoreProperties, PptxCustomProperty } from 'pptx-viewer-core';

	import { useTranslator } from '../../../i18n/context';
	import CustomPropertiesBlock from './CustomPropertiesBlock.svelte';

	const {
		coreProperties,
		appProperties,
		customProperties,
		canEdit = true,
		onupdatecore,
		onupdateapp,
		onupdatecustom,
	}: {
		coreProperties?: PptxCoreProperties;
		appProperties?: PptxAppProperties;
		customProperties: PptxCustomProperty[];
		canEdit?: boolean;
		onupdatecore: (patch: Partial<PptxCoreProperties>) => void;
		onupdateapp: (patch: Partial<PptxAppProperties>) => void;
		onupdatecustom: (next: PptxCustomProperty[]) => void;
	} = $props();
	const t = useTranslator();

	interface FieldRow {
		label: string;
		value: string;
		commit: (value: string) => void;
	}

	const fields = $derived<FieldRow[]>([
		{
			label: t('pptx.properties.titleLabel'),
			value: coreProperties?.title ?? '',
			commit: (value) => onupdatecore({ title: value }),
		},
		{
			label: t('pptx.properties.author'),
			value: coreProperties?.creator ?? '',
			commit: (value) => onupdatecore({ creator: value }),
		},
		{
			label: t('pptx.documentProperties.summary.company'),
			value: appProperties?.company ?? '',
			commit: (value) => onupdateapp({ company: value }),
		},
		{
			label: t('pptx.documentProperties.statistics.application'),
			value: appProperties?.application ?? '',
			commit: (value) => onupdateapp({ application: value }),
		},
	]);
</script>

<div class="pptx-svelte-doc-props">
	{#each fields as field (field.label)}
		<label>
			<span>{field.label}</span>
			<input
				type="text"
				disabled={!canEdit}
				value={field.value}
				onchange={(event) => field.commit(event.currentTarget.value)}
			/>
		</label>
	{/each}
	<CustomPropertiesBlock {customProperties} {canEdit} onupdate={onupdatecustom} />
</div>

<style>
	.pptx-svelte-doc-props {
		display: grid;
		gap: 7px;
	}

	label {
		display: grid;
		gap: 3px;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		font-size: 11px;
	}

	input {
		min-width: 0;
		height: 25px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 5px;
		background: var(--pptx-background, var(--office-background, #ffffff));
		color: inherit;
	}
</style>
