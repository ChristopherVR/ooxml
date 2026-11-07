/** UI adapters only: parsing, editing and serialization stay in the existing engines. */
import { onTheme, currentTheme, setThemeMode } from './themes.js';
import { editorColors, applyOfficeTheme } from './theme-tokens.js';
import { setEditorReadOnly } from './editor-read-only.js';
import { cloudReadOnly } from './cloud-files.js';
import { download } from './ui.js';
import { editorGoogleFonts, googleFontsOption } from './font-downloads.js';
import {
	authorName,
	editorLocale,
	editorRadius,
	externalContentOption,
	onSettings,
	settings,
} from './settings.js';
import { currentProfile } from './profile-state.js';

/** PowerPoint's dictionary for `locale` (English is built in). */
async function pptxMessages(locale) {
	if (!locale || locale === 'en') return undefined;
	const loaders = {
		fr: () => import('pptx-viewer-locales/fr'),
		es: () => import('pptx-viewer-locales/es'),
		de: () => import('pptx-viewer-locales/de'),
		'zh-CN': () => import('pptx-viewer-locales/zh-CN'),
	};
	const module = await loaders[locale]?.();
	return module && Object.values(module).find((value) => value && typeof value === 'object');
}

const author = (prefs) => authorName(currentProfile().name, prefs);

/** Settings > View, Language and General for Word, Excel and Visio (all live). */
function applyEditorSettings(element, kind, prefs, keys) {
	const has = (key) => !keys || keys.includes(key);
	const locale = editorLocale(prefs);
	if (has('locale') && (locale || keys)) {
		// Back to "each app's own choice": English, as the editors start.
		if (kind === 'vsdx') element.setAttribute('locale', locale ?? 'en');
		else element.locale = locale ?? 'en';
	}
	if (has('showRibbon')) element.showToolbar = prefs.showRibbon;
	if (has('authorName')) {
		const name = author(prefs);
		if (kind === 'docx' && (name || keys)) element.reviewAuthor = name ?? '';
		if (kind === 'xlsx' && (name || keys)) element.authorName = name ?? '';
	}
	if (kind === 'docx' && has('showThumbnails')) element.showThumbnails = prefs.showThumbnails;
	if (kind === 'xlsx' && has('showFormulaBar')) element.showFormulaBar = prefs.showFormulaBar;
}
export function preloadEditor(kind) {
	const loaders = {
		pptx: () => import('pptx-vanilla-viewer'),
		docx: () => import('ooxml-ui/docx'),
		xlsx: () => import('ooxml-ui/xlsx'),
		vsdx: () => import('ooxml-ui/visio'),
	};
	return loaders[kind]?.();
}
/** File commands the suite handles itself, for every editor (Visio raises only these). */
export const SUITE_FILE_COMMANDS = new Set(['save', 'new', 'open', 'share']);

export async function mountEditor(host, doc, changed, command) {
	let accessReadOnly = cloudReadOnly(doc);
	if (doc.kind === 'pptx') {
		const { createPptxViewer } = await import('pptx-vanilla-viewer');
		let loading = true;
		let loadFailed = false;
		const googleFonts = googleFontsOption();
		const prefs = settings();
		const locale = editorLocale(prefs);
		// Held by reference by the viewer, so a later language adds its dictionary here.
		const messages = {};
		const dictionary = await pptxMessages(locale);
		if (dictionary) messages[locale] = dictionary;
		const externalContent = externalContentOption(prefs);
		const authorFor = author(prefs);
		const viewer = createPptxViewer(host, {
			// Settings > Fonts; unset keeps the viewer's own Trust Center choice.
			...(googleFonts === undefined ? {} : { googleFonts }),
			...(externalContent === undefined ? {} : { allowExternalContent: externalContent }),
			...(locale ? { locale, defaultLocale: locale } : {}),
			messages,
			...(authorFor ? { authorName: authorFor } : {}),
			showToolbar: prefs.showRibbon,
			showThumbnails: prefs.showThumbnails,
			rulerUnit: prefs.measurementUnit,
			fileName: doc.name,
			editable: !accessReadOnly,
			// The suite owns cloud/Teams sharing; the viewer's separate collaboration invite is not cloud sync.
			hiddenActions: ['share'],
			onOpenFile: () => command('open'),
			storage: {
				async save(bytes, meta) {
					if (meta.reason === 'saveAs') download({ bytes, name: meta.fileName });
					else await command('save');
				},
			},
			onChange: changed,
			autosave: false,
			onError: () => {
				if (loading) loadFailed = true;
			},
			availableThemes: [
				{ key: 'default', labelKey: 'pptx.settings.theme.default', theme: undefined },
			],
			onThemeChange: () => {},
		});
		const stopTheme = onTheme((theme) => {
			applyOfficeTheme(host, theme);
			viewer.setTheme({ colors: editorColors(theme), radius: editorRadius() });
		});
		// The display language changes live; the other options apply to decks opened next.
		const stopSettings = onSettings((next, keys) => {
			if (!keys.includes('locale')) return;
			const code = editorLocale(next) ?? 'en';
			void pptxMessages(code).then((found) => {
				if (found) messages[code] = found;
				viewer.setLocale(code);
			});
		});
		await viewer.loadFile(doc.bytes);
		loading = false;
		if (loadFailed) {
			stopTheme();
			stopSettings();
			viewer.destroy();
			throw new Error('The presentation could not be opened.');
		}
		return {
			save: () => viewer.save(),
			load: async (bytes) => {
				loading = true;
				loadFailed = false;
				try {
					await viewer.loadFile(bytes);
				} finally {
					loading = false;
				}
				if (loadFailed) throw new Error('The presentation could not be opened.');
			},
			readOnly: (on) => viewer.setEditable(!(on || accessReadOnly)),
			setCloudReadOnly: (on) => {
				accessReadOnly = on;
				viewer.setEditable(!on);
			},
			/** Shows where a proposed change lands: its slide, with its shape selected. */
			reveal(change) {
				viewer.goToSlide(change.slideIndex);
				requestAnimationFrame(() => viewer.selectElements([change.elementId]));
			},
			dispose: () => {
				stopTheme();
				stopSettings();
				viewer.destroy();
			},
		};
	}
	let element;
	if (doc.kind === 'docx') {
		(await import('ooxml-ui/docx')).registerDocxEditor();
		element = document.createElement('docx-editor');
	} else if (doc.kind === 'xlsx') {
		(await import('ooxml-ui/xlsx')).defineXlsxEditor();
		element = document.createElement('xlsx-editor');
	} else {
		(await import('ooxml-ui/visio')).registerVisioViewer();
		element = document.createElement('visio-viewer');
	}
	if (doc.kind !== 'vsdx') element.fileName = doc.name;
	// Settings > Fonts: missing fonts download for every app, read when a document loads.
	element.googleFonts = editorGoogleFonts();
	const stopTheme = onTheme((theme) => {
		applyOfficeTheme(element, theme);
		if (doc.kind !== 'vsdx') {
			element.theme = theme.mode;
			element.themeColors = {
				...editorColors(theme),
				...(doc.kind === 'xlsx'
					? { selectionBorder: theme.accent, headerActiveBg: editorColors(theme).accent }
					: {}),
				radius: editorRadius(),
			};
		}
	});
	const observer = new MutationObserver(() => {
		const mode = element.getAttribute('theme');
		if (['light', 'dark'].includes(mode) && mode !== currentTheme().mode) setThemeMode(mode);
	});
	observer.observe(element, { attributes: true, attributeFilter: ['theme'] });
	host.append(element);
	applyEditorSettings(element, doc.kind, settings());
	const stopSettings = onSettings((next, keys) =>
		applyEditorSettings(element, doc.kind, next, keys),
	);
	setEditorReadOnly(element, doc.kind, accessReadOnly);
	// Visio opens at 100%; fit the page to the canvas once it has a layout, as the Visio demo does.
	const fit = () =>
		doc.kind === 'vsdx' && requestAnimationFrame(() => element.isConnected && element.fit());
	await element.load(doc.bytes, doc.name);
	fit();
	for (const name of ['document-change', 'workbook-change'])
		element.addEventListener(name, changed);
	let stop;
	if (doc.kind === 'vsdx') {
		let previous = element.controller.state.edit;
		stop = element.controller.subscribe((state) => {
			if (state.edit !== previous) {
				previous = state.edit;
				changed();
			}
		});
	}
	// Word, Excel and Visio announce New, Open and Save as the cancelable `file-command` event
	// (`ooxml-core/embedding`); the suite takes them over so files stay in its profile storage.
	element.addEventListener('file-command', (event) => {
		if (SUITE_FILE_COMMANDS.has(event.detail.command)) {
			event.preventDefault();
			// Native custom-element events are fire-and-forget; the suite callback reports failures.
			void Promise.resolve()
				.then(() => command(event.detail.command))
				.catch(() => {});
		}
	});
	return {
		save: async () => {
			if (element.commitEdit && !element.commitEdit())
				throw new Error('Finish the current cell edit before saving.');
			return doc.kind === 'vsdx' ? (await element.exportVsdx()).bytes : element.saveBytes();
		},
		load: async (bytes) => {
			await element.load(bytes, doc.name);
			fit();
		},
		clean: () => element.markClean?.(),
		// Puts the cursor in the active cell, as Excel does when a workbook opens.
		focus: () => {
			if (doc.kind === 'xlsx') element.focusGrid();
		},
		// While the assistant previews a change the drawing is read-only (Visio's canEdit is off), so
		// nothing is edited during a review; Discard reloads the stored file and Apply saves the preview.
		readOnly: (on) => setEditorReadOnly(element, doc.kind, on || accessReadOnly),
		setCloudReadOnly: (on) => {
			accessReadOnly = on;
			setEditorReadOnly(element, doc.kind, on);
		},
		/** Shows where a proposed change lands, selected in the editor. Word shows its own redlines. */
		reveal(change) {
			if (change.type === 'cells') {
				element.setActiveSheet(change.sheetIndex);
				element.select(change.cells.map((c) => c.address).join(','));
			}
			if (change.type === 'shape_text') {
				const pages = element.controller.state.document?.pages ?? [];
				const index = pages.findIndex((p) => p.id === change.pageId);
				if (index >= 0) element.pageIndex = index;
				element.selectShapes([{ id: change.shapeId, name: '', pageId: change.pageId }]);
			}
		},
		flush: () => {
			if (element.commitEdit && !element.commitEdit())
				throw new Error('Finish the current cell edit.');
		},
		dispose: () => {
			stop?.();
			stopSettings();
			observer.disconnect();
			stopTheme();
			element.remove();
		},
	};
}

export async function blankDocument(kind) {
	const api = kind === 'pptx' ? null : await import('ooxml-core/automation');
	if (kind === 'docx') return api.createDocx(['']);
	if (kind === 'xlsx') return api.createXlsx(['Sheet1']);
	if (kind === 'pptx') {
		const { PresentationBuilder } = await import('ooxml-core/pptx');
		const { handler, data } = await PresentationBuilder.create({ initialSlideCount: 1 });
		return handler.save(data.slides);
	}
	if (kind === 'vsdx') return (await import('ooxml-core/visio')).createVsdx();
	throw new Error(`Cannot create a blank ${kind} file.`);
}
