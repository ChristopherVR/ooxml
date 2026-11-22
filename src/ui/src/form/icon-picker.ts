import { defineDialog, type OfficeUiDialog } from '../dialog/dialog';
import { defineDialogFooter, type OfficeDialogFooterState } from '../dialog/dialog-footer';
import css from './icon-picker.css?raw';

const ICONS = [
	['Arrow right', 'M4 12h16m-6-6 6 6-6 6'],
	['Bar chart', 'M4 4v16h16M8 16v-4m5 4V7m5 9v-6'],
	['Bell', 'M5 17h14l-2-3V9a5 5 0 0 0-10 0v5zM10 20h4'],
	['Calendar', 'M4 6h16v14H4zM4 10h16M8 4v4m8-4v4'],
	['Camera', 'M3 7h5l2-3h4l2 3h5v13H3zM9 13a3 3 0 1 0 6 0 3 3 0 0 0-6 0z'],
	['Check mark', 'm4 12 5 5L20 6'],
	['Clock', 'M3 12a9 9 0 1 0 18 0 9 9 0 0 0-18 0zm9-5v5l3 2'],
	['Cloud', 'M7 18a4 4 0 1 1 0-8 5 5 0 0 1 9.7 1 3.5 3.5 0 1 1 .8 7z'],
	['Document', 'M6 3h8l4 4v14H6zm8 0v5h4M9 12h6m-6 4h6'],
	['Envelope', 'M3 6h18v12H3zm0 0 9 7 9-7'],
	['Flag', 'M5 21V4c4-3 7 3 14 0v10c-7 3-10-3-14 0'],
	['Heart', 'M12 20 5 13c-6-6 2-12 7-6 5-6 13 0 7 6z'],
	['Home', 'm3 11 9-8 9 8M6 9v12h12V9m-8 12v-6h4v6'],
	['Light bulb', 'M9 17v-3a6 6 0 1 1 6 0v3zm0 3h6'],
	[
		'Location pin',
		'M12 21s-7-6-7-12a7 7 0 0 1 14 0c0 6-7 12-7 12zm0-9a3 3 0 1 0 0-6 3 3 0 0 0 0 6z',
	],
	['Lock', 'M5 10h14v11H5zm3 0V7a4 4 0 1 1 8 0v3m-4 5v2'],
	['Person', 'M8 7a4 4 0 1 0 8 0 4 4 0 0 0-8 0zM4 21v-2a8 6 0 0 1 16 0v2z'],
	['Shield', 'm12 3 8 3v6c0 4-4 7-8 9-4-2-8-5-8-9V6z'],
	['Star', 'm12 3 3 6 7 1-5 5 1 7-6-3-6 3 1-7-5-5 7-1z'],
] as const;

export interface PickedOfficeIcon {
	name: string;
	svg: string;
	path: string;
	color: string;
}

interface IconPickerLabels {
	heading?: string;
	search?: string;
	color?: string;
	empty?: string;
	cancel?: string;
	insert?: string;
	close?: string;
	iconName?: (english: string) => string;
}

/** Opens the bundled icon gallery and returns the chosen graphics in selection order. */
export function pickOfficeIcon(
	host: HTMLElement | ShadowRoot,
	labels: IconPickerLabels = {},
): Promise<PickedOfficeIcon[]> {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	defineDialog(registry);
	defineDialogFooter(registry);
	const dialog = doc.createElement('office-ui-dialog') as OfficeUiDialog;
	dialog.heading = labels.heading ?? 'Icons';
	dialog.closeLabel = labels.close ?? null;
	dialog.style.setProperty('--office-dialog-width', 'min(420px, calc(100vw - 32px))');
	const style = doc.createElement('style');
	style.textContent = css;
	const body = doc.createElement('div');
	body.className = 'office-icon-picker';
	const filters = doc.createElement('div');
	filters.className = 'icon-filters';
	const search = doc.createElement('input');
	search.type = 'search';
	search.placeholder = labels.search ?? 'Search icons';
	search.setAttribute('aria-label', search.placeholder);
	const color = doc.createElement('input');
	color.type = 'color';
	color.value = '#262626';
	color.setAttribute('aria-label', labels.color ?? 'Icon color');
	const grid = doc.createElement('div');
	grid.className = 'icon-grid';
	const empty = doc.createElement('p');
	empty.textContent = labels.empty ?? 'No matching icons.';
	empty.setAttribute('role', 'status');
	const footer = doc.createElement('office-ui-dialog-footer') as HTMLElement & {
		state: OfficeDialogFooterState;
	};
	footer.slot = 'footer';
	filters.append(search, color);
	body.append(filters, grid, empty);
	dialog.append(style, body, footer);

	const selected = new Set<(typeof ICONS)[number]>();
	const nameOf = (name: string) => labels.iconName?.(name) ?? name;
	const svgOf = (path: string, stroke: string, size: number) =>
		`<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24"><path d="${path}" fill="none" stroke="${stroke}" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
	const syncSelection = () => {
		footer.state = {
			actions: [
				{ id: 'cancel', label: labels.cancel ?? 'Cancel' },
				{
					id: 'insert',
					label: labels.insert ?? 'Insert',
					variant: 'primary',
					disabled: selected.size === 0,
				},
			],
		};
	};
	const render = () => {
		const query = search.value.trim().toLocaleLowerCase();
		const matches = ICONS.filter(([name]) =>
			[name, nameOf(name)].some((label) => label.toLocaleLowerCase().includes(query)),
		);
		grid.replaceChildren(
			...matches.map((icon) => {
				const [name, path] = icon;
				const button = doc.createElement('button');
				button.type = 'button';
				button.title = nameOf(name);
				button.setAttribute('aria-label', button.title);
				button.setAttribute('aria-pressed', String(selected.has(icon)));
				// The path is from the bundled catalogue and color is normalized by the color input.
				button.innerHTML = svgOf(path, color.value, 26);
				button.addEventListener('click', () => {
					if (!selected.delete(icon)) selected.add(icon);
					button.setAttribute('aria-pressed', String(selected.has(icon)));
					syncSelection();
				});
				return button;
			}),
		);
		empty.hidden = matches.length > 0;
		syncSelection();
	};
	search.addEventListener('input', render);
	color.addEventListener('input', render);
	render();

	return new Promise((resolve) => {
		dialog.addEventListener('office-dialog-close', () => {
			queueMicrotask(() => dialog.remove());
			resolve([]);
		});
		footer.addEventListener('office-dialog-footer-request', (event) => {
			const { id } = (event as CustomEvent<{ id: string }>).detail;
			if (id === 'insert') {
				resolve(
					[...selected].map(([name, path]) => ({
						name: nameOf(name),
						path,
						color: color.value,
						svg: svgOf(path, color.value, 96),
					})),
				);
			}
			dialog.close();
		});
		host.append(dialog);
		dialog.show();
	});
}
