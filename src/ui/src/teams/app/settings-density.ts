import { html } from 'lit';

export type ChatDensity = 'comfy' | 'compact';

/** Native radios retain keyboard and screen-reader behavior. */
export function densityChoices(density: ChatDensity, change: (density: ChatDensity) => void) {
	return html`<fieldset class="density-options">
		<legend>Chat density</legend>
		${(['comfy', 'compact'] as const).map(
			(value) => html`<label class="density-choice">
				<input
					type="radio"
					name="chat-density"
					value=${value}
					.checked=${density === value}
					@change=${() => change(value)}
				/>
				<span
					><strong>${value === 'comfy' ? 'Comfy' : 'Compact'}</strong>
					<small
						>${value === 'comfy' ? 'More space between messages' : 'More messages in less space'}</small
					></span
				>
			</label>`,
		)}
	</fieldset>`;
}
