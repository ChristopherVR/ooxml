/**
 * The Settings panes the shared options rows cannot express: the appearance cards, the fonts
 * choice, the privacy overview and the assistant summary. Each returns an element for one slot
 * of `office-ui-options-dialog`; `settings-dialog.js` puts them in and wires their events.
 */
import { escape } from './ui.js';
import { accents, themePreferences } from './themes.js';
import { fontDownloads, normalizeFontServer } from './font-downloads.js';
import { settings } from './settings.js';

const REPO = 'https://github.com/ChristopherVR/ooxml/blob/main';

function pane(slot, html) {
	const element = document.createElement('div');
	element.slot = slot;
	element.className = 'settings-pane';
	element.innerHTML = html;
	return element;
}

const radio = (name, value, on) =>
	`<input type="radio" name="${name}" value="${value}" ${on ? 'checked' : ''} />`;

/** Light, dark or system cards, then the accent swatches. */
export function appearanceBlock(slot) {
	const { mode, accent } = themePreferences();
	const modes = [
		['light', 'Light'],
		['dark', 'Dark'],
		['system', 'Use system setting'],
	];
	return pane(
		slot,
		`<fieldset class="settings-modes"><legend>Theme</legend>${modes
			.map(
				([value, label]) =>
					`<label>${radio('mode', value, mode === value)}<span class="theme-preview ${value}" aria-hidden="true"><i></i><i></i><i></i></span><span>${label}</span></label>`,
			)
			.join(
				'',
			)}</fieldset><fieldset class="settings-accents"><legend>Accent colour</legend>${Object.entries(
			accents,
		)
			.map(
				([value, [name, light, dark]]) =>
					`<label title="${escape(name)}">${radio('accent', value, accent === value)}<span class="settings-swatch" style="--swatch:${light};--swatch-dark:${dark}" aria-hidden="true"></span><span>${escape(name)}</span></label>`,
			)
			.join('')}</fieldset>`,
	);
}

/** What the current font choice contacts, for the privacy overview. */
export function fontDownloadsSummary(choice = fontDownloads()) {
	if (choice.mode === 'off') return { on: false, text: 'Off: no font requests leave this device.' };
	if (choice.mode === 'custom') {
		const host = (() => {
			try {
				return new URL(choice.server).host;
			} catch {
				return choice.server;
			}
		})();
		return { on: true, text: `From your font server, ${host}.` };
	}
	return { on: true, text: 'From Google Fonts (fonts.googleapis.com, fonts.gstatic.com).' };
}

export function fontsPane(slot) {
	const choice = fontDownloads();
	const option = (value, title, text, extra = '') =>
		`<div class="settings-choice"><label>${radio('fonts', value, choice.mode === value)}<span><strong>${title}</strong><small>${text}</small></span></label>${extra}</div>`;
	return pane(
		slot,
		`<section><h3>Fonts that are missing</h3><p class="settings-desc">Documents often name fonts that are neither installed on this device nor embedded in the file, such as Aptos and other Microsoft 365 cloud fonts. OOXML Office can download the same font, or one with identical letter widths (Carlito for Calibri), so text, line breaks and pages come out as they do in Office.</p><div class="settings-choices" role="radiogroup" aria-label="Fonts that are missing">${option(
			'google',
			'Download from Google Fonts',
			"Recommended. Your browser asks Google for the families the document needs, which shares your IP address with Google. The file is never sent, and fonts outside Google's public catalogue are never named.",
			'<p class="settings-hosts"><code>fonts.googleapis.com</code><code>fonts.gstatic.com</code></p>',
		)}${option(
			'custom',
			'Download from my own font server',
			'A server you or your organisation runs that answers in the Google Fonts CSS2 format. Nothing goes to Google.',
			`<div class="settings-field"><input type="url" name="font-server" placeholder="https://fonts.example.com/css2" aria-label="Font server address" autocomplete="url" spellcheck="false" value="${escape(choice.server)}" /><p class="settings-error" role="status" aria-live="polite"></p></div>`,
		)}${option(
			'off',
			'Do not download fonts',
			'No font requests leave this device. Text uses installed, embedded or similar fonts, so line and page breaks can differ from Office.',
		)}</div><p class="settings-note">Applies to documents you open next, in Word, Excel, PowerPoint and Visio.</p></section><section><h3>Where this applies</h3><dl class="settings-apps"><div><dt>Word, Excel and Visio</dt><dd>Follow this setting.</dd></div><div><dt>PowerPoint</dt><dd>Follows this setting. With Google Fonts chosen, its own File &gt; Options &gt; Trust Center switch can still turn downloads off.</dd></div><div><dt>Fonts in the file</dt><dd>Fonts embedded in a presentation are used first and never downloaded.</dd></div></dl></section><section><h3>Self-hosting</h3><p class="settings-desc">This site's Content Security Policy allows only Google's two font hosts, so a deployment that uses its own server adds that server to the policy. <a href="${REPO}/docs/font-loading.md" target="_blank" rel="noopener">How font loading works</a></p></section>`,
	);
}

/** The custom server field's message, or '' when the address is usable. */
export function fontServerError(value) {
	if (!value.trim()) return 'Enter the address of your font server.';
	return normalizeFontServer(value)
		? ''
		: 'Enter a full http:// or https:// address without a sign-in, query (?) or fragment (#).';
}

const connection = (on, title, text, goto) =>
	`<li data-state="${on ? 'on' : 'off'}"><span class="settings-dot" aria-hidden="true"></span><span><strong>${title}</strong><small>${escape(text)}</small></span>${goto ? `<button type="button" data-goto="${goto}">Change</button>` : ''}</li>`;

/** `assistant`: `null` when the policy turns it off, else `{ offline, provider }`. */
export function privacyPane(slot, { assistant, cloud }) {
	const fonts = fontDownloadsSummary();
	const blocked = settings().blockExternalContent;
	return pane(
		slot,
		`<section><h3>Connections</h3><p class="settings-desc">OOXML Office has no server of its own and collects no analytics or telemetry. Opening, editing and saving happen in this browser. It connects elsewhere only for the following.</p><ul class="settings-connections">${[
			connection(fonts.on, 'Missing fonts', fonts.text, 'fonts'),
			connection(
				!blocked,
				'Linked pictures and media',
				blocked
					? 'Blocked in presentations you open.'
					: "PowerPoint's Trust Center decides (allowed by default).",
			),
			assistant
				? connection(
						!assistant.offline,
						'AI assistant',
						assistant.offline
							? 'Offline preview: nothing leaves this device.'
							: `Sends your prompts and the documents you attach to ${assistant.provider}.`,
						'assistant',
					)
				: connection(false, 'AI assistant', 'Turned off on this site.'),
			connection(
				Boolean(cloud),
				'Cloud storage',
				cloud
					? `Files you open from ${cloud}, and its sign-in, go to that service.`
					: 'Not set up on this site.',
			),
			connection(
				false,
				'Sharing and Teams',
				'Only when you share a document, start a live session or set up a Teams server.',
			),
			connection(
				true,
				'App updates',
				'Checks this site for a newer version of the app. No other server.',
			),
		].join(
			'',
		)}</ul><p class="settings-note">Every connection also sends what any web request sends: your IP address and browser details. Text boxes use your browser's spell checker; some browsers' enhanced spell check sends text to the browser's maker.</p></section><section><h3>On this device</h3><p class="settings-desc">Your documents, versions, attachments and chat history are kept in this browser for each profile, and your settings in its local storage. Nothing here is uploaded. Clearing this site's data in your browser removes it.</p><div class="settings-storage"><span data-storage>Checking storage...</span><button type="button" data-persist hidden>Keep my files when space runs low</button></div></section><p class="settings-links"><a href="${REPO}/PRIVACY.md" target="_blank" rel="noopener">Read the full privacy statement</a></p>`,
	);
}

const size = (bytes) =>
	bytes >= 1e9
		? `${(bytes / 1e9).toFixed(1)} GB`
		: bytes >= 1e6
			? `${(bytes / 1e6).toFixed(1)} MB`
			: `${Math.max(1, Math.round(bytes / 1e3))} KB`;

/** Fill the privacy pane's storage line from the browser's own estimate. */
export async function describeStorage(element) {
	const line = element.querySelector('[data-storage]');
	const persist = element.querySelector('[data-persist]');
	const storage = globalThis.navigator?.storage;
	if (!line) return;
	try {
		const { usage = 0, quota = 0 } = (await storage?.estimate?.()) ?? {};
		const kept = (await storage?.persisted?.()) ?? false;
		line.textContent = `${size(usage)} used${quota ? ` of ${size(quota)} available` : ''}. ${kept ? 'The browser keeps these files even when space runs low.' : 'The browser may remove them if the device runs low on space.'}`;
		if (persist) persist.hidden = kept || typeof storage?.persist !== 'function';
	} catch {
		line.textContent = 'This browser does not report how much space the app uses.';
	}
}

export function assistantPane(slot, { label, managed }) {
	return pane(
		slot,
		`<section><h3>Model</h3><p class="settings-desc">${managed ? 'Your organisation manages the models and most assistant settings. ' : ''}The assistant uses <strong>${escape(label)}</strong>. Choose another model from the menu under the assistant's message box.</p></section><section><h3>Replies and behaviour</h3><p class="settings-desc">Reply length, your own instructions, which documents are attached, and the models and keys you added.</p><p><button type="button" data-assistant>Assistant settings...</button></p></section>`,
	);
}
