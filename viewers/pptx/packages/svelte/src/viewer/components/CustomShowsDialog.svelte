<script lang="ts">
	import { createCustomShowsController } from 'ooxml-ui/pptx';
	import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
	import { untrack } from 'svelte';

	import { useTranslator } from '../../i18n/context';
	import { controllerState } from '../state/controller-state';
	import DialogFooter from './DialogFooter.svelte';

	const {
		shows,
		slides,
		activeShowId = null,
		onclose,
		onsave,
		onsetactive,
		onrun,
	}: {
		shows: PptxCustomShow[];
		slides: PptxSlide[];
		/** The show playback is currently restricted to, or null for the whole deck. */
		activeShowId?: string | null;
		onclose: () => void;
		/** Commit the shows; `clearedSlides` is set when jumps into removed shows were cleared. */
		onsave: (shows: PptxCustomShow[], clearedSlides: PptxSlide[] | undefined) => void;
		/** Restrict (or, with null, stop restricting) the slide show to one show. */
		onsetactive?: (id: string | null) => void;
		/** Start the show once the dialog has committed and selected it. */
		onrun?: (show: PptxCustomShow) => void;
	} = $props();
	const t = useTranslator();

	// All the dialog's logic is the shared controller; this file only draws it.
	const controller = createCustomShowsController(
		untrack(() => shows),
		untrack(() => activeShowId),
	);
	const snapshot = controllerState(controller);
	const view = $derived.by(() => {
		void snapshot.current;
		return controller.view(slides);
	});
	/** The picker's value. '' means "All Slides", i.e. no restriction. */
	let active = $state<string>(untrack(() => activeShowId) ?? '');
	const pickerValue = $derived(view.shows.some(({ id }) => id === active) ? active : '');
	function setActive(next: string): void {
		active = next;
		onsetactive?.(next || null);
	}
	function commit(): PptxCustomShow[] {
		const result = controller.result(slides);
		onsave(result.shows, result.slides);
		// A show removed in this session must not leave playback pinned to an id nothing resolves.
		onsetactive?.(result.shows.some(({ id }) => id === active) ? active : null);
		return result.shows;
	}
	function ok(): void {
		commit();
		onclose();
	}
	function run(): void {
		const id = controller.getSnapshot().selectedId;
		const show = commit().find((candidate) => candidate.id === id);
		if (!show) {
			return;
		}
		onsetactive?.(show.id);
		onrun?.(show);
		onclose();
	}
</script>

<div class="backdrop">
	<section role="dialog" tabindex="-1" aria-modal="true" aria-label={t('pptx.customShows.title')}>
		<header>
			<h2>{t('pptx.customShows.title')}</h2>
			<button onclick={onclose} aria-label={t('pptx.customShows.close')}>×</button>
		</header>
		<div class="active-picker">
			<span>{t('pptx.customShows.selectCustomShow')}</span>
			<pptx-ui-select
				aria-label={t('pptx.customShows.selectCustomShow')}
				value={pickerValue}
				onchange={(event) => setActive(event.currentTarget.value)}
			>
				<option value="">{t('pptx.customShows.allSlides')}</option>
				{#each view.shows as show (show.id)}<option value={show.id}>{show.name}</option>{/each}
			</pptx-ui-select>
		</div>
		<div class="body">
			<aside>
				<ul role="listbox" aria-label={t('pptx.customShows.savedShows')}>
					{#each view.shows as show (show.id)}
						<li
							role="option"
							aria-selected={show.selected}
							tabindex="0"
							onclick={() => controller.select(show.id)}
							onkeydown={(event) => event.key === 'Enter' && controller.select(show.id)}
						>
							{show.name}
						</li>
					{:else}
						<li class="empty">{t('pptx.customShows.empty')}</li>
					{/each}
				</ul>
				<button onclick={() => controller.create(t('pptx.customShows.newDefaultName'))}>
					{t('pptx.customShows.createNew')}
				</button>
				<button disabled={!view.canEdit} onclick={() => controller.copySelected()}>
					{t('pptx.customShows.copy')}
				</button>
				<button disabled={!view.canEdit} onclick={() => controller.removeSelected()}>
					{t('pptx.customShows.remove')}
				</button>
			</aside>
			<main>
				{#if view.selected}
					<label class="name">
						<span>{t('pptx.customShows.renameNameLabel')}</span>
						<input
							aria-label={t('pptx.customShows.renameNameLabel')}
							value={view.selected.name}
							onchange={(event) => controller.rename(event.currentTarget.value)}
						/>
					</label>
					<div class="lists">
						<div>
							<strong>{t('pptx.customShows.slidesInPresentation')}</strong>
							<ul role="listbox" aria-label={t('pptx.customShows.slidesInPresentation')} aria-multiselectable="true">
								{#each view.presentationSlides as row (row.rId)}
									<li
										role="option"
										aria-selected={row.selected}
										tabindex="0"
										onclick={() => controller.toggleLeft(row.rId)}
										ondblclick={() => controller.addSlide(row.rId)}
										onkeydown={(event) => event.key === ' ' && controller.toggleLeft(row.rId)}
									>
										{row.number}. {row.title}
									</li>
								{/each}
							</ul>
						</div>
						<div class="actions">
							<button disabled={!view.canAdd} onclick={() => controller.addSelected(slides)}>
								{t('pptx.customShows.add')}
							</button>
							<button disabled={!view.canRemoveSlides} onclick={() => controller.removeSelectedSlides()}>
								{t('pptx.customShows.remove')}
							</button>
						</div>
						<div>
							<strong>{t('pptx.customShows.slidesInShow')}</strong>
							<ul role="listbox" aria-label={t('pptx.customShows.slidesInShow')} aria-multiselectable="true">
								{#each view.showSlides as row (row.position)}
									<li
										role="option"
										aria-selected={row.selected}
										tabindex="0"
										onclick={() => controller.toggleRight(row.position ?? 0)}
										onkeydown={(event) => event.key === ' ' && controller.toggleRight(row.position ?? 0)}
									>
										{row.number}. {row.title}
									</li>
								{:else}
									<li class="empty">{t('pptx.customShows.noSlidesYet')}</li>
								{/each}
							</ul>
						</div>
						<div class="actions">
							<button disabled={!view.canMoveUp} onclick={() => controller.moveUp()}>
								{t('pptx.customShows.moveUp')}
							</button>
							<button disabled={!view.canMoveDown} onclick={() => controller.moveDown()}>
								{t('pptx.customShows.moveDown')}
							</button>
						</div>
					</div>
				{:else}
					<p>{t('pptx.customShows.noSlidesYet')}</p>
				{/if}
			</main>
		</div>
		<footer>
			<DialogFooter
				actions={[
					{ id: 'show', label: t('pptx.customShows.showButton') },
					{ id: 'cancel', label: t('pptx.customShows.cancel') },
					{ id: 'save', label: t('pptx.common.ok'), variant: 'primary' },
				]}
				onaction={(id) => (id === 'save' ? ok() : id === 'show' ? run() : onclose())}
			/>
		</footer>
	</section>
</div>

<style>
	.backdrop {
		position: fixed;
		inset: 0;
		z-index: 1200;
		display: grid;
		place-items: center;
		background: #0009;
	}
	section {
		width: min(820px, calc(100vw - 32px));
		max-height: 88vh;
		overflow: auto;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 12px;
		background: var(--pptx-card, var(--office-surface, #f3f4f6));
	}
	header,
	footer {
		display: flex;
		align-items: center;
		justify-content: space-between;
		padding: 13px 17px;
		border-bottom: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
	}
	footer {
		justify-content: flex-end;
		border-top: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-bottom: 0;
	}
	h2,
	p {
		margin: 0;
	}
	h2 {
		font-size: 14px;
	}
	.active-picker {
		display: flex;
		align-items: center;
		gap: 8px;
		padding: 10px 17px;
		border-bottom: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		font-size: 12px;
	}
	.body {
		display: grid;
		grid-template-columns: 200px 1fr;
		min-height: 320px;
	}
	aside {
		display: flex;
		flex-direction: column;
		gap: 6px;
		padding: 12px;
		border-right: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
	}
	main {
		display: grid;
		align-content: start;
		gap: 10px;
		padding: 14px;
	}
	.lists {
		display: grid;
		grid-template-columns: 1fr auto 1fr auto;
		gap: 8px;
	}
	.actions {
		display: flex;
		flex-direction: column;
		justify-content: center;
		gap: 6px;
	}
	ul {
		margin: 0;
		padding: 0;
		max-height: 240px;
		overflow: auto;
		list-style: none;
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 4px;
	}
	li {
		padding: 5px 8px;
		cursor: pointer;
		font-size: 12px;
	}
	li[aria-selected='true'] {
		background: var(--pptx-primary, #2563eb);
		color: #fff;
	}
	li.empty {
		cursor: default;
		color: var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));
	}
	button,
	input {
		border: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		border-radius: 6px;
		padding: 7px;
		background: var(--pptx-muted, var(--office-surface, #f3f4f6));
		color: inherit;
	}
	button:disabled {
		opacity: 0.5;
	}
	header button {
		border: 0;
		background: transparent;
		font-size: 20px;
	}
	.name {
		display: flex;
		align-items: center;
		gap: 8px;
		font-size: 12px;
	}
	.name input {
		flex: 1;
	}
	@media (max-width: 600px) {
		section {
			position: fixed;
			inset: auto 0 0;
			width: 100%;
			max-height: 88dvh;
		}
		.body,
		.lists {
			grid-template-columns: 1fr;
		}
		aside {
			border-right: 0;
			border-bottom: 1px solid var(--pptx-border, var(--office-border, #d1d5db));
		}
	}
</style>
