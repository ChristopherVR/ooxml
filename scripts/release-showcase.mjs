import { highlight } from './release-highlight.mjs';
import { escapeHtml } from './release-markdown.mjs';
import { renderSecurityScene } from './release-security-scenes.mjs';

/** The release a showcase belongs to; 2.0's configuration predates the field. */
const versionOf = (config) => config?.version ?? '2.0';

export function renderReleaseIndex(sections, version = '2.0') {
	if (!sections) return '';
	return `<nav class="release-index" aria-label="${escapeHtml(version)} highlights">${sections.map(({ id, label }) => `<a href="#${escapeHtml(id)}">${escapeHtml(label)}</a>`).join('')}</nav>`;
}

function renderMetrics({ action, context, start, before, after }, version) {
	const max = Math.max(before.value ?? 0, after.value ?? 0);
	const comparable = before.value !== undefined;
	const reading = (version, metric) => {
		const delay = comparable ? Math.round((metric.value / max) * 2500) : metric.value ? 730 : 160;
		const progress = comparable
			? `<span class="benchmark-track" aria-hidden="true"><i style="--duration:${delay}ms;--measure:${metric.value / max}"></i></span>`
			: '';
		return `<div class="benchmark-reading">
			<dt>${version}</dt><dd style="--duration:${delay}ms">${escapeHtml(metric.label)}</dd>${progress}
		</div>`;
	};
	return `<div class="showcase-copy benchmark-copy">
		<h3>${escapeHtml(action)}</h3>
		<p>${escapeHtml(context)}</p>
		<div class="benchmark-start"><span aria-hidden="true"></span>${escapeHtml(start)}</div>
		<dl class="benchmark-readings">${reading(`Before ${version}`, before)}${reading(`OOXML ${version}`, after)}</dl>
		${comparable ? '<p class="benchmark-note">Animation scaled to show relative time.</p>' : ''}
	</div>`;
}

function renderFlow({ title, text, keys, steps }) {
	const shortcut = keys
		?.map(
			(key, i) =>
				`<button type="button" data-step-direction="${i ? -1 : 1}" aria-label="${i ? 'Previous' : 'Next'} region (${escapeHtml(key)})"><kbd>${escapeHtml(key)}</kbd></button>`,
		)
		.join('');
	return `<div class="showcase-copy flow-copy">
		${shortcut ? `<div class="flow-shortcut">${shortcut}</div>` : ''}
		<h3>${escapeHtml(title)}</h3>
		<p>${escapeHtml(text)}</p>
		<ol class="flow-steps">${steps.map((step, i) => `<li style="--step:${i}"${keys && i === 0 ? ' aria-current="step"' : ''}><i aria-hidden="true"></i><span>${escapeHtml(step)}</span></li>`).join('')}</ol>
	</div>`;
}

function picture({ src, width, height, animation }, alt, eager = false) {
	const size = width && height ? ` width="${width}" height="${height}"` : '';
	const gif = animation
		? ` data-animation="${escapeHtml(animation)}" data-poster="${escapeHtml(src)}"`
		: '';
	return `<img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}"${gif}${size}${eager ? ' fetchpriority="high"' : ' loading="lazy"'} decoding="async" />`;
}

function renderShell(product, media, eager) {
	const image = media(`assets/shell-${product.product}.webp`);
	return `<div class="showcase-preview">${picture(image, product.alt, eager)}</div>
		<div class="install-command">
			<pre><span aria-hidden="true">$ </span><code>${highlight(product.command, 'sh')}</code></pre>
			<button type="button" data-copy-command="${escapeHtml(product.command)}">Copy</button>
		</div>`;
}

function renderCapture(media, file, alt, after = false) {
	const image = media(`assets/${file}`);
	const alternate = after ? ' data-capture-after' : '';
	return `<button class="showcase-capture" type="button" aria-haspopup="dialog" data-preview${alternate} aria-label="Enlarge preview: ${escapeHtml(alt)}">${picture(image, alt)}<span aria-hidden="true">Enlarge ↗</span></button>`;
}

function renderFeature(item, media) {
	const preview = renderCapture(media, item.image, item.text);
	return `<div class="feature-caption"><h3>${escapeHtml(item.title)}</h3><p>${escapeHtml(item.text)}</p></div>
		<div class="showcase-preview">${preview}</div>`;
}

const labels = {
	shell: 'Install OOXML',
	performance: 'Performance comparisons',
	compatibility: 'Safari and iPadOS improvements',
	security: 'Security improvements',
	accessibility: 'Accessibility improvements',
	'new-features': 'New editor features',
};

/** Products whose tile a tab shows even when the showcase has no install tour. */
const PRODUCT_TILES = new Set(['word', 'excel', 'powerpoint', 'visio', 'openteams']);

/** A workspace, benchmark, feature capture or flow with direct selection. */
export function renderShowcase(config, type, { media, tile, notes = [] }) {
	const items = config?.[type];
	if (!items?.length || !(type in labels)) return '';
	const version = versionOf(config);
	const label = type === 'shell' ? `${labels.shell} ${version}` : labels[type];
	const products = new Map((config.shell ?? []).map((product) => [product.product, product]));
	const panels = items.map((item, i) => {
		const product = item.product ? ` data-product="${escapeHtml(item.product)}"` : '';
		let content;
		if (type === 'shell') content = renderShell(item, media, !i);
		else if (type === 'performance') content = renderMetrics(item, version);
		else if (type === 'new-features') content = renderFeature(item, media);
		else if (type === 'security')
			content = renderSecurityScene(item, (file, alt, after) =>
				renderCapture(media, file, alt, after),
			);
		else {
			const flow = renderFlow(item);
			if (item.capture) {
				const preview = renderCapture(media, item.capture, item.captureAlt ?? item.title);
				content = `<div class="flow-lens">${flow}<div class="showcase-preview">${preview}</div></div>`;
			} else content = flow;
		}
		return `<div class="showcase-panel"${product}${i ? ' hidden' : ''}>${content}</div>`;
	});
	const choices = items.map((item, i) => {
		const product = products.get(item.product);
		const area = item.product === 'teams' ? 'openteams' : item.product;
		const icon = product || PRODUCT_TILES.has(area) ? tile(area) : '';
		return `<button type="button" data-slide title="Select to keep this tab open" aria-controls="showcase-${type}-stage" aria-pressed="${i === 0}">${icon}<span>${escapeHtml(item.label ?? product.label)}</span></button>`;
	});
	const id = type === 'shell' ? ' id="install"' : '';
	const details = notes.length
		? `<div class="showcase-notes">${notes.map((text, i) => `<div class="showcase-note" data-note="${i}"${i ? ' hidden' : ''}>${text}</div>`).join('')}</div>`
		: '';
	return `<div${id} class="showcase showcase-${type}" aria-label="${label}" role="region" data-showcase="${type}">
		<div class="showcase-controls" hidden>${choices.join('')}</div>
		<div class="showcase-stage" id="showcase-${type}-stage">${panels.join('')}</div>${details}
	</div>`;
}
