import { escapeHtml } from './release-markdown.mjs';

const intro = ({ title, text }) => `<h3>${escapeHtml(title)}</h3><p>${escapeHtml(text)}</p>`;

function renderConsent(item, capture) {
	return `<div class="flow-lens" data-security-scene="consent">
		<div class="showcase-copy security-copy" data-gated>
			${intro(item)}
			<div class="consent-state"><span data-consent-blocked>External media blocked</span><span data-consent-allowed>External media allowed</span></div>
			<button class="flow-trigger" type="button" data-unlock>Enable content</button>
		</div>
		<div class="showcase-preview">
			${capture(item.capture, item.captureAlt)}
			${capture(item.captureAfter, item.afterAlt, true)}
		</div>
	</div>`;
}

function renderToolOutput({ request, output }) {
	return `<div class="security-tool">
		<div><code>xlsx_set_cells</code><span>A2</span></div>
		<pre class="security-request"><code>${escapeHtml(request)}</code></pre>
		<pre><code>${escapeHtml(output)}</code></pre>
	</div>`;
}

export function renderSecurityScene(item, capture) {
	if (item.scene === 'consent') return renderConsent(item, capture);
	const scene = escapeHtml(item.scene);
	const choices = item.examples.map(
		({ label }, i) =>
			`<button type="button" data-security-choice="${i}" aria-controls="security-${scene}-${i}" aria-pressed="${i === 0}">${escapeHtml(label)}</button>`,
	);
	const examples = item.examples.map((example, i) => {
		const visual = example.capture
			? capture(example.capture, example.captureAlt)
			: renderToolOutput(example);
		return `<div class="security-case" id="security-${scene}-${i}" data-security-case="${i}"${i ? ' hidden' : ''}>
			<div class="security-result" data-tone="${escapeHtml(example.tone ?? 'neutral')}">
				<strong>${escapeHtml(example.status)}</strong><span>${escapeHtml(example.description)}</span>
			</div>
			${visual}
		</div>`;
	});
	return `<div class="flow-lens security-scene" data-security-scene="${scene}">
		<div class="showcase-copy security-copy">
			${intro(item)}
			<div class="security-choices" role="group" aria-label="Compare examples">${choices.join('')}</div>
		</div>
		<div class="security-cases">${examples.join('')}</div>
	</div>`;
}
