/** UI adapters only: parsing, editing and serialization stay in the existing engines. */
import { onTheme, currentTheme, setThemeMode } from './themes.js';
import { editorColors, applyOfficeTheme } from './theme-tokens.js';
export function preloadEditor(kind) {
	const loaders = {
		pptx: () => import('pptx-vanilla-viewer'),
		docx: () => import('ooxml-ui/docx'),
		xlsx: () => import('ooxml-ui/xlsx'),
		vsdx: () => import('ooxml-ui/visio'),
	};
	return loaders[kind]?.();
}
export async function mountEditor(host, doc, changed, command) {
	if (doc.kind === 'pptx') {
		const { createPptxViewer } = await import('pptx-vanilla-viewer');
		const viewer = createPptxViewer(host, {
			fileName: doc.name,
			editable: true,
			onChange: changed,
			autosave: false,
			availableThemes: [
				{ key: 'default', labelKey: 'pptx.settings.theme.default', theme: undefined },
			],
			onThemeChange: () => {},
		});
		const stopTheme = onTheme((theme) => {
			applyOfficeTheme(host, theme);
			viewer.setTheme({ colors: editorColors(theme), radius: '0.25rem' });
		});
		await viewer.loadFile(doc.bytes);
		return {
			save: () => viewer.save(),
			load: (bytes) => viewer.loadFile(bytes),
			dispose: () => {
				stopTheme();
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
	const stopTheme = onTheme((theme) => {
		applyOfficeTheme(element, theme);
		if (doc.kind !== 'vsdx') {
			element.theme = theme.mode;
			element.themeColors = {
				...editorColors(theme),
				...(doc.kind === 'xlsx'
					? { selectionBorder: theme.accent, headerActiveBg: editorColors(theme).accent }
					: {}),
				radius: '4px',
			};
		}
	});
	const observer = new MutationObserver(() => {
		const mode = element.getAttribute('theme');
		if (['light', 'dark'].includes(mode) && mode !== currentTheme().mode) setThemeMode(mode);
	});
	observer.observe(element, { attributes: true, attributeFilter: ['theme'] });
	host.append(element);
	await element.load(doc.bytes, doc.name);
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
	element.addEventListener('file-command', (event) => {
		if (['save', 'new', 'open'].includes(event.detail.command)) {
			event.preventDefault();
			command(event.detail.command);
		}
	});
	return {
		save: async () => {
			if (element.commitEdit && !element.commitEdit())
				throw new Error('Finish the current cell edit before saving.');
			return doc.kind === 'vsdx' ? (await element.exportVsdx()).bytes : element.saveBytes();
		},
		load: (bytes) => element.load(bytes, doc.name),
		clean: () => element.markClean?.(),
		flush: () => {
			if (element.commitEdit && !element.commitEdit())
				throw new Error('Finish the current cell edit.');
		},
		dispose: () => {
			stop?.();
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
