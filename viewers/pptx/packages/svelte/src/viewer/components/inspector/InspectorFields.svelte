<script lang="ts" generics="T, P">
	/**
	 * InspectorFields: the Svelte binding's one generic field renderer. It draws
	 * any list of shared inspector section descriptors (a toggle per section,
	 * then a two-column grid of its fields); values, bounds, labels and patches
	 * all come from the descriptors, so this component only maps them onto
	 * markup. A colour field previews while its native picker is dragged
	 * (`onpreview`, no undo step) and commits once when the pick settles
	 * (`onpatch`, then `oncolorcommit` for the "Recent colours" list), through
	 * the shared `createLiveColorPatchInput`.
	 */
	import {
		createLiveColorPatchInput,
		inspectorFieldPatch,
		inspectorTogglePatch,
		resolveInspectorSections,
	} from 'ooxml-ui/pptx';
	import type {
		InspectorFieldValue,
		InspectorSectionDescriptor,
		LiveColorInput,
	} from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import { liveColorEvents } from './live-color';

	const {
		sections,
		target,
		onpatch,
		onpreview,
		oncolorcommit,
		disabled = false,
	}: {
		sections: readonly InspectorSectionDescriptor<T, P>[];
		target: T;
		onpatch: (patch: P) => void;
		/** A colour drag's live patch: apply it without an undo step. */
		onpreview?: (patch: P) => void;
		oncolorcommit?: (hex: string) => void;
		disabled?: boolean;
	} = $props();
	const t = useTranslator();

	const views = $derived(resolveInspectorSections(sections, target));

	function onField(sectionId: string, fieldId: string, raw: InspectorFieldValue): void {
		const patch = inspectorFieldPatch(sections, sectionId, fieldId, target, raw);
		if (patch !== undefined) {
			onpatch(patch);
		}
	}

	/** One live colour controller per colour field, created on first use. */
	const colorInputs = new Map<string, LiveColorInput>();

	function colorInput(sectionId: string, fieldId: string): LiveColorInput {
		const key = `${sectionId}.${fieldId}`;
		let control = colorInputs.get(key);
		if (!control) {
			control = createLiveColorPatchInput<T, P & object>({
				read: () => {
					const field = views
						.find((section) => section.id === sectionId)
						?.fields.find((candidate) => candidate.id === fieldId);
					return field && { key, color: String(field.value), state: target };
				},
				patch: (value, hex) =>
					inspectorFieldPatch(sections, sectionId, fieldId, value, hex) as
						| (P & object)
						| undefined,
				preview: (patch) => onpreview?.(patch),
				commit: (patch) => onpatch(patch),
				remember: (hex) => oncolorcommit?.(hex),
			});
			colorInputs.set(key, control);
		}
		return control;
	}

	function onToggle(sectionId: string, on: boolean): void {
		const patch = inspectorTogglePatch(sections, sectionId, target, on);
		if (patch !== undefined) {
			onpatch(patch);
		}
	}
</script>

<div class="pptx-svelte-fields">
	{#each views as section (section.id)}
		{#if section.toggleable}
			<!-- svelte-ignore a11y_label_has_associated_control -- form-associated custom element -->
			<label class="pptx-svelte-fields-toggle">
				<pptx-ui-checkbox
					data-testid={section.toggleTestId}
					checked={section.on}
					{disabled}
					onchange={(e) => onToggle(section.id, e.currentTarget.checked)}
				></pptx-ui-checkbox>
				<span>{t(section.labelKey ?? '')}</span>
			</label>
		{:else if section.labelKey}
			<span class="pptx-svelte-fields-title">{t(section.labelKey)}</span>
		{/if}
		{#if section.fields.length}
			<div class="pptx-svelte-fields-grid" data-section={section.id}>
				{#each section.fields as field (field.id)}
					{#if field.control === 'checkbox'}
						<!-- svelte-ignore a11y_label_has_associated_control -- form-associated custom element -->
						<label class="pptx-svelte-fields-check">
							<pptx-ui-checkbox
								data-testid={field.testId}
								checked={Boolean(field.value)}
								{disabled}
								onchange={(e) => onField(section.id, field.id, e.currentTarget.checked)}
							></pptx-ui-checkbox>
							<span>{t(field.labelKey)}</span>
						</label>
					{:else}
						<!-- svelte-ignore a11y_label_has_associated_control -- form-associated custom element -->
						<label class:pptx-svelte-fields-wide={field.wide}>
							{t(field.labelKey)}
							{#if field.control === 'select'}
								<pptx-ui-select
									data-testid={field.testId}
									aria-label={t(field.labelKey)}
									value={String(field.value)}
									{disabled}
									onchange={(e) => onField(section.id, field.id, e.currentTarget.value)}
								>
									{#each field.options ?? [] as option (option.value)}
										<option value={option.value}>{t(option.labelKey)}</option>
									{/each}
								</pptx-ui-select>
							{:else if field.control === 'color'}
								<input
									type="color"
									data-testid={field.testId}
									aria-label={t(field.labelKey)}
									value={String(field.value)}
									{disabled}
									use:liveColorEvents={colorInput(section.id, field.id)}
								/>
							{:else if field.control === 'range'}
								<input
									type="range"
									data-testid={field.testId}
									aria-label={t(field.labelKey)}
									min={field.min}
									max={field.max}
									step={field.step}
									value={Number(field.value)}
									{disabled}
									oninput={(e) => onField(section.id, field.id, e.currentTarget.value)}
								/>
							{:else}
								<input
									type="number"
									data-testid={field.testId}
									aria-label={t(field.labelKey)}
									min={field.min}
									max={field.max}
									step={field.step}
									value={Number(field.value)}
									{disabled}
									onchange={(e) => onField(section.id, field.id, e.currentTarget.value)}
								/>
							{/if}
						</label>
					{/if}
				{/each}
			</div>
		{/if}
	{/each}
</div>

<style>
	.pptx-svelte-fields {
		display: flex;
		flex-direction: column;
		gap: 4px;
		margin-top: 6px;
	}

	.pptx-svelte-fields-toggle {
		display: flex;
		align-items: center;
		gap: 6px;
		margin-top: 8px;
		cursor: pointer;
	}

	.pptx-svelte-fields-title {
		margin-top: 8px;
		font-weight: 600;
	}

	.pptx-svelte-fields-grid {
		display: grid;
		grid-template-columns: 1fr 1fr;
		gap: 6px;
		margin-left: 4px;
	}

	.pptx-svelte-fields-grid label {
		display: grid;
		gap: 3px;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
		font-size: 10px;
	}

	.pptx-svelte-fields-grid .pptx-svelte-fields-wide {
		grid-column: 1 / -1;
	}

	.pptx-svelte-fields-check {
		grid-column: 1 / -1;
		display: flex !important;
		flex-direction: row !important;
		align-items: center;
		gap: 6px;
	}

	.pptx-svelte-fields-grid input {
		height: 26px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 5px;
		background: var(--pptx-background, var(--office-background, #ffffff));
		color: inherit;
	}
</style>
