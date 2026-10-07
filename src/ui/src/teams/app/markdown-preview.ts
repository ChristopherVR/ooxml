import { html, nothing } from 'lit';
import { markdownBlocks, markdownInline } from 'ooxml-core/teams';

function inline(text: string, base: string) {
	return markdownInline(text, base).map((token) => {
		if (token.kind === 'strong') return html`<strong>${token.text}</strong>`;
		if (token.kind === 'emphasis') return html`<em>${token.text}</em>`;
		if (token.kind === 'code') return html`<code>${token.text}</code>`;
		if (token.kind === 'link')
			return html`<a href=${token.url!} target="_blank" rel="noopener noreferrer"
				>${token.text}</a
			>`;
		return token.text;
	});
}

/** Render core tokens as Lit text bindings. Never insert peer-provided HTML. */
export function markdownPreview(source: string, base: string) {
	return markdownBlocks(source).map((block) => {
		if (block.kind === 'code') return html`<pre><code>${block.text}</code></pre>`;
		if (block.kind === 'table' && block.table) {
			const table = block.table;
			return html`<div
				class="markdown-table"
				tabindex="0"
				role="region"
				aria-label="Markdown table"
			>
				<table>
					<thead>
						<tr>
							${table.headers.map((cell, column) => html`<th scope="col" data-align=${table.align[column] ?? 'left'}>${inline(cell, base)}</th>`)}
						</tr>
					</thead>
					<tbody>
						${table.rows.map(
							(row) =>
								html`<tr>
									${row.map((cell, column) => html`<td data-align=${table.align[column] ?? 'left'}>${inline(cell, base)}</td>`)}
								</tr>`,
						)}
					</tbody>
				</table>
			</div>`;
		}
		const text = inline(block.text, base);
		if (block.kind === 'quote') return html`<blockquote>${text}</blockquote>`;
		if (block.kind === 'list')
			return html`<ul class=${block.checked === undefined ? '' : 'task-list'}>
				<li>
					${block.checked === undefined ? nothing : html`<input type="checkbox" aria-label=${block.text} ?checked=${block.checked} disabled />`}${text}
				</li>
			</ul>`;
		if (block.kind === 'heading')
			return html`<div role="heading" aria-level=${block.level} class="heading">${text}</div>`;
		return html`<p>${text}</p>`;
	});
}
