/**
 * Settings: the shared `office-ui-options-dialog` (the element PowerPoint's File > Options uses)
 * filled with the suite's categories. Every edit applies at once, as in PowerPoint; Cancel puts
 * back what was set when the dialog opened, and "Reset all settings" restores the defaults,
 * appearance and fonts included.
 */
import { defineOptionsDialog } from 'ooxml-ui/controls';
import { notify } from './ui.js';
import { DEFAULT_THEME, onTheme, setThemePreferences, themePreferences } from './themes.js';
import { fontDownloads, normalizeFontServer, setFontDownloads } from './font-downloads.js';
import { resetSettings, settings, updateSettings } from './settings.js';
import { currentProfile } from './profile-state.js';
import { assistantEnabled, isManaged } from './ai/policy.js';
import { connectionLabel, currentConnection, provider } from './ai/providers.js';
import {
	blockSlot,
	settingsCategories,
	settingsDefaults,
	settingsLabels,
	settingsValues,
} from './settings-model.js';
import {
	appearanceBlock,
	assistantPane,
	describeStorage,
	fontServerError,
	fontsPane,
	privacyPane,
} from './settings-panes.js';

let dialog;

const cloudLabel = () =>
	globalThis.ooxmlCloudOptions?.label ?? (globalThis.ooxmlOneDriveOptions ? 'OneDrive' : '');

function assistantState() {
	if (!assistantEnabled()) return null;
	const c = currentConnection();
	return { offline: !c || c.provider === 'offline', provider: c ? provider(c.provider).name : '' };
}

/** Put the panes of the open category in their slots. */
function renderSlots() {
	for (const child of [...dialog.children]) child.remove();
	switch (dialog.category) {
		case 'general':
			dialog.append(appearanceBlock(blockSlot('appearance')));
			break;
		case 'fonts':
			dialog.append(fontsPane('pane-fonts'));
			break;
		case 'privacy': {
			const element = privacyPane('pane-privacy', {
				assistant: assistantState(),
				cloud: cloudLabel(),
			});
			dialog.append(element);
			void describeStorage(element);
			break;
		}
		case 'assistant':
			dialog.append(
				assistantPane('pane-assistant', { label: connectionLabel(), managed: isManaged() }),
			);
			break;
	}
}

/** The slotted panes' own inputs: theme cards, font choice and server, buttons. */
function onPaneChange(event) {
	const target = event.target;
	if (!(target instanceof HTMLInputElement) || !dialog.contains(target)) return;
	if (target.name === 'mode' || target.name === 'accent')
		setThemePreferences({ [target.name]: target.value });
	if (target.name === 'fonts' || target.name === 'font-server') applyFontChoice();
}

function applyFontChoice() {
	const mode = dialog.querySelector('input[name=fonts]:checked')?.value;
	const field = dialog.querySelector('input[name=font-server]');
	const error = dialog.querySelector('.settings-error');
	const custom = mode === 'custom';
	if (field) field.closest('.settings-choice').dataset.active = String(custom);
	if (error) error.textContent = '';
	if (!custom) {
		setFontDownloads({ mode });
		return;
	}
	const message = fontServerError(field?.value ?? '');
	// Until the address is usable the previous choice holds; nothing falls back to Google.
	if (message) {
		if (error && field.value.trim()) error.textContent = message;
		field?.focus();
		return;
	}
	setFontDownloads({ mode, server: normalizeFontServer(field.value) });
}

function onPaneClick(event) {
	const button = event.target.closest?.('button[type=button]');
	if (!button || !dialog.contains(button)) return;
	if (button.dataset.goto) {
		dialog.category = button.dataset.goto;
		dialog.shadowRoot?.querySelector('nav [aria-selected="true"]')?.focus();
	}
	if (button.hasAttribute('data-persist'))
		void navigator.storage
			.persist()
			.then(() => describeStorage(button.closest('.settings-pane')))
			.catch(notify);
	if (button.hasAttribute('data-assistant')) {
		dialog.open = false;
		void import('./ai/settings.js').then(({ openPreferences }) => openPreferences()).catch(notify);
	}
}

function create() {
	defineOptionsDialog();
	dialog = document.createElement('office-ui-options-dialog');
	dialog.id = 'settings-dialog';
	dialog.heading = 'Settings';
	dialog.live = true;
	dialog.labels = settingsLabels;
	dialog.defaults = settingsDefaults();
	let snapshot;
	dialog.addEventListener('office-option-input', (event) => {
		updateSettings({ [event.detail.key]: event.detail.value });
		// The privacy overview reflects the linked-content switch beside it.
		if (event.detail.key === 'blockExternalContent') renderSlots();
	});
	dialog.addEventListener('office-options-category', renderSlots);
	dialog.addEventListener('office-options-reset', () => {
		resetSettings();
		setThemePreferences(DEFAULT_THEME);
		setFontDownloads({ mode: 'google' });
		dialog.values = settingsValues(settings());
		renderSlots();
	});
	dialog.addEventListener('office-options-cancel', () => {
		if (!snapshot) return;
		updateSettings(snapshot.settings);
		setThemePreferences(snapshot.theme);
		setFontDownloads(snapshot.fonts);
	});
	dialog.addEventListener('change', onPaneChange);
	dialog.addEventListener('input', (event) => {
		if (event.target?.name === 'font-server') applyFontChoice();
	});
	dialog.addEventListener('click', onPaneClick);
	dialog.snapshot = () => {
		snapshot = { settings: settings(), theme: themePreferences(), fonts: fontDownloads() };
	};
	document.body.append(dialog);
	// The element picks its native control scheme when it opens; a live theme change updates it.
	onTheme((theme) => {
		dialog.style.colorScheme = theme.mode;
	});
}

/** Open Settings, on `category` when given. */
export function openSettings(category) {
	if (!dialog) create();
	dialog.categories = settingsCategories({
		profileName: currentProfile().name ?? '',
		assistant: assistantEnabled(),
	});
	dialog.values = settingsValues(settings());
	if (category) dialog.category = category;
	dialog.snapshot();
	dialog.show();
	renderSlots();
}
