<script lang="ts">
	/**
	 * EquationTemplateGallery: the equation editor's starter-formula tiles.
	 * Split out of `EquationEditorDialog` to keep that file within the repo's
	 * file-size budget.
	 *
	 * The catalogue (`EQUATION_TEMPLATES`) is shared across every binding; each
	 * tile's MathML is compiled once here, on mount, because the set is static.
	 * Every rendered string goes through `sanitizeMathMl` (inside shared's
	 * `compileEquationTemplateMathMl`) before the `{@html ...}` binding.
	 */
	import { compileEquationTemplateMathMl, EQUATION_TEMPLATES } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../../i18n/context';

	const {
		activeLatex,
		onselect,
	}: {
		/** The editor's current LaTeX; the matching tile shows as active. */
		activeLatex: string;
		onselect: (latex: string) => void;
	} = $props();
	const t = useTranslator();

	const TEMPLATE_MATHML = compileEquationTemplateMathMl();
</script>

<div class="templates">
	<span class="label">{t('pptx.equation.templates')}</span>
	<div class="grid">
		{#each EQUATION_TEMPLATES as tmpl, idx (tmpl.latex)}
			<button
				type="button"
				class="template"
				class:active={activeLatex === tmpl.latex}
				title={t(tmpl.i18nKey)}
				onclick={() => onselect(tmpl.latex)}
			>
				<!-- eslint-disable-next-line svelte/no-at-html-tags -->
				<span class="math tile">{@html TEMPLATE_MATHML[idx]}</span>
				<span class="template-label">{t(tmpl.i18nKey)}</span>
			</button>
		{/each}
	</div>
</div>

<style>
	.templates {
		display: flex;
		flex-direction: column;
		gap: 5px;
	}

	.label {
		font-size: 11px;
		font-weight: 500;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}

	.math {
		font-family: 'Cambria Math', 'STIX Two Math', serif;
	}

	.grid {
		display: grid;
		grid-template-columns: repeat(4, minmax(0, 1fr));
		gap: 6px;
	}

	.template {
		display: flex;
		flex-direction: column;
		align-items: center;
		gap: 4px;
		padding: 7px;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: var(--pptx-radius, var(--office-radius, 4px));
		background: var(--pptx-muted, var(--office-surface, #f3f4f6));
		color: inherit;
		cursor: pointer;
	}

	.template:hover {
		background: var(--pptx-accent, var(--office-selected, #e5e7eb));
	}

	.template.active {
		border-color: var(--pptx-primary, var(--office-accent, #2563eb));
		background: color-mix(in srgb, var(--pptx-primary, var(--office-accent, #2563eb)) 14%, transparent);
	}

	.tile {
		display: flex;
		align-items: center;
		justify-content: center;
		height: 28px;
		overflow: hidden;
		font-size: 13px;
	}

	.template-label {
		width: 100%;
		overflow: hidden;
		text-overflow: ellipsis;
		white-space: nowrap;
		text-align: center;
		font-size: 9px;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}
</style>
