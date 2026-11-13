import { computed, DestroyRef, effect, inject, Injectable, signal, untracked } from '@angular/core';
import type { Signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	createFindBarController,
	createPptxBuildingBlockSession,
	registerBuildingBlockGroups,
	resolveBuildingBlockMode,
	runBuildingBlockToolbarControl,
} from 'ooxml-ui/pptx';
import type { BuildingBlockToolbarControlId, ViewerBuildingBlockActions } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';

import { controllerSignal } from '../viewer/controller-signal';
import { ListConversionScopeService } from '../viewer/list-conversion-scope.service';
import { LoadNoticesService } from '../viewer/load-notices.service';
import { POWER_POINT_VIEWER_PROVIDERS } from '../viewer/power-point-viewer.providers';
import { ViewerCollaborationShellService } from '../viewer/viewer-collaboration-shell.service';
import { ViewerCustomShowsService } from '../viewer/viewer-custom-shows.service';
import { ViewerDialogsService } from '../viewer/viewer-dialogs.service';
import { ViewerDocumentPropertiesService } from '../viewer/viewer-document-properties.service';
import { ViewerFindReplaceService } from '../viewer/viewer-find-replace.service';
import { ViewerKeyboardService } from '../viewer/viewer-keyboard.service';
import { ViewerZoomService } from '../viewer/viewer-zoom.service';
import { createAngularGroupPorts } from './group-ports';
import { createAngularPort } from './session-port';
import type { ViewerBuildingBlocksOptions, ViewerBuildingBlocksState } from './types';
import { routeDialogFlag } from './route-dialog-flag';

const bytesOf = (source: ViewerBuildingBlocksOptions['content']) =>
	source instanceof ArrayBuffer || source instanceof Uint8Array ? source : null;

/**
 * The Angular building blocks: the shared building-blocks session (`ooxml-ui/pptx`) over the
 * Angular editor services. Provide it per host component with {@link VIEWER_BUILDING_BLOCKS_PROVIDERS},
 * feed it options with `createViewerBuildingBlocks(() => options)` (or `update`), and place
 * `pptx-building-blocks-canvas`, `pptx-building-blocks-toolbar` and `pptx-slides-pane` with `[blocks]`.
 * The blocks mount no dialogs or pickers: controls and canvas gestures that need one raise a
 * request (`onRequest`, `actions().pending`) the host finishes with `actions()`.
 */
@Injectable()
export class ViewerBuildingBlocks {
	readonly shell = inject(ViewerCollaborationShellService);
	readonly zoom = inject(ViewerZoomService);
	private readonly dialogs = inject(ViewerDialogsService);
	private readonly docProps = inject(ViewerDocumentPropertiesService);
	private readonly customShows = inject(ViewerCustomShowsService);
	private readonly findService = inject(ViewerFindReplaceService);
	private readonly keyboard = inject(ViewerKeyboardService);
	private readonly notices = inject(LoadNoticesService);
	private readonly translation = inject(TranslateService, { optional: true });
	private readonly destroyRef = inject(DestroyRef);

	private readonly optionsSignal = signal<ViewerBuildingBlocksOptions>({});
	private readonly canvas = signal<{
		host: HTMLElement;
		stage: () => HTMLElement | undefined;
	} | null>(null);
	private announced = false;

	readonly translate = (key: string): string => {
		const found = this.translation?.instant(key);
		return found && found !== key
			? found
			: ((translationsEn as Record<string, string>)[key] ?? key);
	};

	private readonly findBar = createFindBarController({
		slides: () => this.shell.editor.slides(),
		commit: (slides, scope) =>
			this.shell.editor.applyReplacement(
				[...slides],
				this.translate(scope === 'all' ? 'pptx.undoAction.replaceAll' : 'pptx.undoAction.replace'),
			),
		navigate: ({ slideIndex, elementId }) => {
			this.shell.activeSlideIndex.set(slideIndex);
			this.shell.editor.select(elementId ? [elementId] : []);
		},
	});
	/** The host find bar's open flag and session, as a signal. */
	readonly findBarState = controllerSignal(this.findBar, this.destroyRef);

	private readonly session = createPptxBuildingBlockSession(
		createAngularPort({
			shell: this.shell,
			notices: this.notices,
			options: () => this.optionsSignal(),
			findBar: this.findBar,
			canvasHost: () => this.canvas()?.host ?? null,
			onCancel: () => this.setShortcutsOpen(false),
		}),
	);
	private readonly pending = controllerSignal(this.session.requests, this.destroyRef);

	/** The host-owned workflows; a new object whenever anything they report changes. */
	readonly actions: Signal<ViewerBuildingBlockActions> = computed(() => {
		this.pending();
		this.findBarState();
		this.optionsSignal();
		return this.session.actions();
	});

	readonly state: Signal<ViewerBuildingBlocksState> = computed(() => {
		const { shell } = this;
		const count = shell.editor.slides().length;
		return {
			mode: resolveBuildingBlockMode('edit', shell.canEdit()),
			loading: shell.loader.loading() || shell.sourcePending(),
			error: shell.loader.error(),
			slideIndex: count ? shell.activeSlideIndex() : -1,
			slideCount: count,
			canUndo: shell.editor.canUndo(),
			canRedo: shell.editor.canRedo(),
			dirty: shell.editor.dirty(),
		};
	});
	readonly mode = computed(() => this.state().mode);
	readonly loading = computed(() => this.state().loading);
	readonly error = computed(() => this.state().error);
	readonly options: Signal<ViewerBuildingBlocksOptions> = this.optionsSignal.asReadonly();

	constructor() {
		const { shell, dialogs } = this;
		const options = this.optionsSignal;
		// Computeds, so a new options object with the same content never reloads the deck.
		const content = computed(() => options().content ?? null);
		const collaboration = computed(() => options().collaboration);
		const canEdit = computed(() => options().canEdit === true);
		shell.bind({
			content,
			collaboration,
			canEdit,
			stageElement: () => this.canvas()?.stage(),
			initialSlide: () => options().initialSlide,
		});
		shell.loader.bindHostLimits(() => options().limits);
		shell.loader.bindHostGoogleFonts(() => options().googleFonts);
		shell.loader.bindHostAllowExternalContent(() => options().allowExternalContent);
		inject(ListConversionScopeService).bind(computed(() => options().listConversionScope));
		this.customShows.bind({
			activeSlideIndex: shell.activeSlideIndex,
			liveSlides: () => shell.editor.slides(),
			applySlides: (slides, label) => shell.editor.applyReplacement(slides, label),
		});
		shell.editor.bindSlideNavigation((index) => this.goTo(index));
		this.keyboard.bind({
			canEdit: shell.canEdit,
			presenting: () => false,
			activeSlideIndex: shell.activeSlideIndex,
			goPrev: () => this.goTo(shell.activeSlideIndex() - 1),
			goNext: () => this.goTo(shell.activeSlideIndex() + 1),
			isEditingText: () => shell.canvasEditing.editingId() !== null,
			selectedElement: () => {
				const id = shell.editor.selectedIds()[0];
				return id ? (shell.editor.elementById(shell.activeSlideIndex(), id) ?? null) : null;
			},
		});
		registerBuildingBlockGroups(
			this.session,
			createAngularGroupPorts({
				shell,
				dialogs,
				customShows: this.customShows,
				notices: this.notices,
				content: () => bytesOf(options().content),
				setShortcutsOpen: (open) => this.setShortcutsOpen(open),
				translate: this.translate,
			}),
		);
		// The editor opens its own dialogs by flag; nothing renders them here, so each becomes a request.
		routeDialogFlag(
			() => this.docProps.showHyperlink(),
			() => {
				this.docProps.showHyperlink.set(false);
				this.session.request({ kind: 'hyperlink' });
			},
		);
		routeDialogFlag(
			() => dialogs.showEquation(),
			() => {
				const existingOmml = dialogs.editingEquationOmml();
				dialogs.showEquation.set(false);
				dialogs.editingEquationOmml.set(null);
				dialogs.editingEquationElementId.set(null);
				this.session.request({ kind: 'equation', existingOmml });
			},
		);
		// Ctrl+F toggles the host find bar, Ctrl+H opens it.
		effect(() => {
			const find = this.findService.showFind();
			const replace = this.findService.showFindReplace();
			if (find || replace) {
				untracked(() => {
					this.findService.showFind.set(false);
					this.findService.showFindReplace.set(false);
					if (find) {
						this.findBar.toggle();
					} else {
						this.findBar.setOpen(true);
					}
				});
			}
		});
		// `?`, Ctrl+/ and Escape change the shared flag; mirror it as a request.
		effect(() => {
			const open = dialogs.showShortcuts();
			untracked(() => this.announceShortcuts(open));
		});
		let lastDirty = false;
		effect(() => {
			const dirty = shell.editor.dirty();
			untracked(() => {
				if (dirty !== lastDirty) {
					lastDirty = dirty;
					options().onDirtyChange?.(dirty);
				}
			});
		});
	}

	/** Apply new options: every one applies at once except `limits` and `initialSlide` (next load). */
	update(next: ViewerBuildingBlocksOptions): void {
		this.optionsSignal.set(next);
	}

	/** Show a slide and clear the selection (as the editor does on a slide change). */
	goTo(index: number): void {
		const count = this.shell.editor.slides().length;
		if (index >= 0 && index < count && index !== this.shell.activeSlideIndex()) {
			this.shell.activeSlideIndex.set(index);
			this.shell.editor.clearSelection();
		}
	}

	setShortcutsOpen(open: boolean): void {
		this.dialogs.showShortcuts.set(open);
		this.announceShortcuts(open);
	}

	private announceShortcuts(open: boolean): void {
		if (open !== this.announced) {
			this.announced = open;
			this.session.announce({ kind: 'shortcuts', open }, open);
		}
	}

	/** Run a toolbar control (the toolbar block calls this). */
	runToolbar(id: BuildingBlockToolbarControlId): void {
		runBuildingBlockToolbarControl(id, this.session, {
			undo: () => this.shell.editor.undo(),
			redo: () => this.shell.editor.redo(),
			toggleFind: () => this.findBar.toggle(),
			zoomOut: () => this.zoom.zoomOut(),
			zoomIn: () => this.zoom.zoomIn(),
			zoomToFit: () => this.zoom.zoomReset(),
			previousSlide: () => this.goTo(this.shell.activeSlideIndex() - 1),
			nextSlide: () => this.goTo(this.shell.activeSlideIndex() + 1),
		});
	}

	/** Called by the canvas block while it is mounted. */
	attachCanvas(host: HTMLElement, stage: () => HTMLElement | undefined): () => void {
		const value = { host, stage };
		this.canvas.set(value);
		return () => {
			if (this.canvas() === value) {
				this.canvas.set(null);
			}
		};
	}

	/** Forward a keydown to the shared editor keymap (the canvas block calls this). */
	handleKeyDown(event: KeyboardEvent): void {
		const host = this.canvas()?.host;
		const target = event.target as Node | null;
		const inCanvas = Boolean(host && target && host.contains(target));
		const onPage = target === document.body || target === document.documentElement;
		// F5 would start a show the blocks do not mount; other chords outside the canvas are the host's.
		if (event.key === 'F5' || !(inCanvas || onPage)) {
			return;
		}
		this.keyboard.handleKeyDown(event);
	}
}

/** The providers a host component lists to own one building-blocks composition. */
export const VIEWER_BUILDING_BLOCKS_PROVIDERS = [
	...POWER_POINT_VIEWER_PROVIDERS,
	ViewerCollaborationShellService,
	ViewerBuildingBlocks,
] as const;

/** Create the blocks in an injection context and keep their options in step with `options()`. */
export function createViewerBuildingBlocks(
	options: () => ViewerBuildingBlocksOptions = () => ({}),
): ViewerBuildingBlocks {
	const blocks = inject(ViewerBuildingBlocks);
	untracked(() => blocks.update(options()));
	effect(() => {
		const next = options();
		untracked(() => blocks.update(next));
	});
	return blocks;
}
