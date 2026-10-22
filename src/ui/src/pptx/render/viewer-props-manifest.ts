/**
 * The shared PowerPoint viewer surface, as one manifest.
 *
 * All five bindings (react, vue, angular, svelte, vanilla) ship the same viewer, and their
 * public props, inputs, options and events used to drift one binding at a time. This file names
 * every concept once. `viewers/pptx/scripts/check-viewer-props.mjs` reads each binding's public
 * type from source and fails when a binding exposes something the manifest does not list, or
 * lacks something the manifest expects, unless the manifest records the exception and why.
 *
 * Naming follows each framework: a concept has a canonical camelCase id, and a binding spells
 * it by rule (events) or by an explicit override in `names` (props). Idiomatic differences
 * stay (`content` against `source`, `canEdit` against `editable`, `className` against `class`);
 * a concept a binding does not have goes in `absent` with the reason, so a missing prop is
 * always a decision on record rather than an accident.
 *
 * Framework-neutral and dependency-free on purpose: the check loads it as plain TypeScript.
 */

export const VIEWER_BINDINGS = ['react', 'vue', 'angular', 'svelte', 'vanilla'] as const;
export type ViewerBinding = (typeof VIEWER_BINDINGS)[number];

export interface ViewerSurfaceEntry {
	/** Canonical camelCase name of the concept. */
	readonly id: string;
	/** How a binding spells it when that is not the default for its kind. */
	readonly names?: Readonly<Partial<Record<ViewerBinding, string>>>;
	/** Bindings that intentionally lack it, with the reason. */
	readonly absent?: Readonly<Partial<Record<ViewerBinding, string>>>;
}

/** Every binding outside `present` lacks the concept, for `reason`. */
const onlyIn = (
	present: readonly ViewerBinding[],
	reason: string,
): Partial<Record<ViewerBinding, string>> =>
	Object.fromEntries(
		VIEWER_BINDINGS.filter((binding) => !present.includes(binding)).map((binding) => [
			binding,
			reason,
		]),
	);

const SOURCE_NAMES = { svelte: 'source', vanilla: 'source' } as const;
const EDITABLE_NAMES = { svelte: 'editable', vanilla: 'editable' } as const;
const CLASS_NAMES = { vue: 'class', angular: 'class', svelte: 'class' } as const;

const BINDING_ONLY = 'binding-only option with no equivalent in the other bindings yet';

/** Props, inputs and options. Callbacks are events (see {@link VIEWER_EVENTS}). */
export const VIEWER_PROPS: readonly ViewerSurfaceEntry[] = [
	{ id: 'content', names: SOURCE_NAMES },
	{ id: 'fonts' },
	{ id: 'fitPadding' },
	{ id: 'maxFitScale' },
	{ id: 'filePath' },
	{ id: 'fileName' },
	{ id: 'canEdit', names: EDITABLE_NAMES },
	{ id: 'autosave' },
	{ id: 'autosaveIntervalMs' },
	{
		id: 'className',
		names: CLASS_NAMES,
		absent: { vanilla: 'the host owns the container element it passes to createPptxViewer' },
	},
	{ id: 'authorName' },
	{ id: 'theme' },
	{ id: 'defaultThemeKey' },
	{ id: 'availableThemes' },
	{ id: 'defaultLocale' },
	{ id: 'availableLocales' },
	{ id: 'accountAuth' },
	{ id: 'collaboration' },
	{ id: 'shareDefaults' },
	{ id: 'smartArt3D' },
	{ id: 'surfaceChart3D' },
	{ id: 'barChart3D' },
	{ id: 'lineChart3D' },
	{ id: 'areaChart3D' },
	{ id: 'pieChart3D' },
	{ id: 'initialSlide' },
	{ id: 'showToolbar' },
	{ id: 'showThumbnails' },
	{ id: 'showCompatibilityToasts' },
	{ id: 'hiddenActions' },
	{ id: 'ribbonAddIns' },
	{ id: 'inspectorInitiallyOpen' },
	{ id: 'customization' },
	{ id: 'ai' },
	// Binding-only conveniences, kept for compatibility. The deprecated ones name their
	// replacement in the binding's own type documentation.
	{
		id: 'locale',
		absent: onlyIn(
			['svelte', 'vanilla'],
			'defaultLocale is the initial language here; locale is a svelte/vanilla extra',
		),
	},
	{
		id: 'messages',
		absent: onlyIn(['vanilla'], 'translations come from the framework i18n in this binding'),
	},
	{
		id: 'showNotes',
		absent: onlyIn(['svelte'], 'deprecated svelte-only flag; use customization.hiddenPanels'),
	},
	{
		id: 'readOnly',
		absent: onlyIn(['vanilla'], 'deprecated vanilla-only flag; use editable'),
	},
	{
		id: 'showInspector',
		absent: onlyIn(['vanilla'], 'deprecated vanilla-only flag; use customization.hiddenPanels'),
	},
	{
		id: 'autosaveFilePath',
		absent: onlyIn(['vanilla'], 'deprecated vanilla-only alias of filePath'),
	},
	{ id: 'showFormatToolbar', absent: onlyIn(['vanilla'], BINDING_ONLY) },
	{ id: 'registry', absent: onlyIn(['vanilla'], BINDING_ONLY) },
	{ id: 'rulerUnit', absent: onlyIn(['vanilla'], BINDING_ONLY) },
];

/**
 * Events. Default spelling by binding: react `onX`, vue `x-y` emit, angular `xY` output,
 * svelte `onxy`, vanilla `onX`. Vue and Angular take some host callbacks as props or inputs
 * (`onHyperlinkClick`, `onOpenFile`, `onThemeChange`, `onLocaleChange`); those are named in
 * `names`.
 */
export const VIEWER_EVENTS: readonly ViewerSurfaceEntry[] = [
	{ id: 'dirtyChange' },
	{ id: 'contentChange' },
	{
		id: 'activeSlideChange',
		names: { svelte: 'onslidechange', vanilla: 'onSlideChange' },
	},
	{ id: 'modeChange' },
	{ id: 'zoomChange' },
	{ id: 'selectionChange' },
	{ id: 'slideCountChange' },
	{ id: 'load' },
	{ id: 'error' },
	{ id: 'startCollaboration' },
	{ id: 'stopCollaboration' },
	// The shared `CollaborationState` every Office editor reports (`collaboration-change`).
	{ id: 'collaborationChange' },
	{ id: 'hyperlinkClick', names: { vue: 'onHyperlinkClick', angular: 'onHyperlinkClick' } },
	{ id: 'openFile', names: { vue: 'onOpenFile', angular: 'onOpenFile' } },
	{
		id: 'themeChange',
		names: { vue: 'onThemeChange', angular: 'onThemeChange', svelte: 'onThemeChange' },
	},
	{
		id: 'localeChange',
		names: { vue: 'onLocaleChange', angular: 'onLocaleChange', svelte: 'onLocaleChange' },
	},
	// The host-facing autosave and edit events, in every binding. Two Angular outputs are spelled
	// differently on purpose: `autosave` is an input there, and an output named `change` would
	// also receive the native `change` events of the inputs inside the viewer.
	{ id: 'autosave', names: { angular: 'autosaved' } },
	{ id: 'autosaveToggle', names: { vanilla: 'onToggleAutosave' } },
	{ id: 'propertiesChange' },
	{ id: 'change', names: { angular: 'documentChange' } },
	{ id: 'notesUpdate' },
	{ id: 'presentationChange' },
	{ id: 'autosaveStatus' },
	{ id: 'autosaveRecovery' },
	// Raw connection status only; `collaborationChange` carries the full shared state everywhere.
	{ id: 'collaborationStatus' },
];

const capitalised = (id: string): string => id.charAt(0).toUpperCase() + id.slice(1);

/** The name a binding gives a prop: its override, else the canonical id. */
export function viewerPropName(entry: ViewerSurfaceEntry, binding: ViewerBinding): string {
	return entry.names?.[binding] ?? entry.id;
}

/** The name a binding gives an event: its override, else the framework's rule. */
export function viewerEventName(entry: ViewerSurfaceEntry, binding: ViewerBinding): string {
	const override = entry.names?.[binding];
	if (override) {
		return override;
	}
	switch (binding) {
		case 'react':
		case 'vanilla':
			return `on${capitalised(entry.id)}`;
		case 'vue':
			return entry.id.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);
		case 'angular':
			return entry.id;
		case 'svelte':
			return `on${entry.id.toLowerCase()}`;
	}
}

export interface ViewerSurfaceDrift {
	binding: ViewerBinding;
	kind: 'prop' | 'event';
	/** What is wrong, in a sentence naming the binding and the member. */
	message: string;
}

/** What one binding publicly exposes, as read from its source (raw names, as spelled). */
export interface ExtractedSurface {
	props: readonly string[];
	events: readonly string[];
}

function diffKind(
	kind: 'prop' | 'event',
	entries: readonly ViewerSurfaceEntry[],
	nameOf: (entry: ViewerSurfaceEntry, binding: ViewerBinding) => string,
	extracted: Readonly<Record<ViewerBinding, ExtractedSurface>>,
): ViewerSurfaceDrift[] {
	const drift: ViewerSurfaceDrift[] = [];
	const label = kind === 'prop' ? 'prop/input/option' : 'event';
	for (const binding of VIEWER_BINDINGS) {
		const actual = new Set(extracted[binding][kind === 'prop' ? 'props' : 'events']);
		const known = new Set<string>();
		for (const entry of entries) {
			const name = nameOf(entry, binding);
			const reason = entry.absent?.[binding];
			known.add(name);
			if (reason === undefined && !actual.has(name)) {
				drift.push({
					binding,
					kind,
					message: `${binding} lacks the ${label} "${name}" (manifest id "${entry.id}"); add it, or list ${binding} under its \`absent\` with a reason`,
				});
			} else if (reason !== undefined && actual.has(name)) {
				drift.push({
					binding,
					kind,
					message: `${binding} exposes the ${label} "${name}" although the manifest lists it as absent (${reason}); remove the exception`,
				});
			}
		}
		for (const name of actual) {
			if (!known.has(name)) {
				drift.push({
					binding,
					kind,
					message: `${binding} exposes the ${label} "${name}" which is not in the viewer props manifest; add it there (and to the other bindings, or mark them absent)`,
				});
			}
		}
	}
	return drift;
}

/** Every difference between the manifest and what the bindings expose; empty when in step. */
export function diffViewerSurface(
	extracted: Readonly<Record<ViewerBinding, ExtractedSurface>>,
): ViewerSurfaceDrift[] {
	return [
		...diffKind('prop', VIEWER_PROPS, viewerPropName, extracted),
		...diffKind('event', VIEWER_EVENTS, viewerEventName, extracted),
	];
}
