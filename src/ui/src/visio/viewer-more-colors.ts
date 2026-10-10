import { normalizePaletteHex } from 'ooxml-core/color';
import { ViewerDialog } from './viewer-dialog';

const CHANNELS = ['Red', 'Green', 'Blue'] as const;

/**
 * Office's More Colors dialog, Custom tab: a hex field and Red, Green and Blue (0-255) that
 * follow each other, with New and Current swatches. OK hands the colour (`#rrggbb`) to the
 * caller. The colour wheel and the HSL model are not offered.
 */
export class ViewerMoreColors {
	readonly dialog: ViewerDialog;
	#hex: HTMLInputElement;
	#channels: HTMLInputElement[];
	#next: HTMLElement;
	#current: HTMLElement;
	#pick: ((color: string) => void) | undefined;
	constructor(root: ShadowRoot) {
		const doc = root.ownerDocument;
		this.dialog = new ViewerDialog(
			root,
			'more-colors-dialog',
			'Colors',
			['OK', 'Cancel'],
			(name) => (name === 'OK' ? this.#accept() : this.dialog.close()),
		);
		const field = (text: string, input: HTMLInputElement) => {
			const label = doc.createElement('label');
			label.append(text, input);
			return label;
		};
		this.#hex = doc.createElement('input');
		this.#hex.type = 'text';
		this.#hex.maxLength = 7;
		this.#hex.spellcheck = false;
		this.#hex.dataset.colorField = 'hex';
		this.#hex.placeholder = '#RRGGBB';
		this.#channels = CHANNELS.map((name) => {
			const input = doc.createElement('input');
			input.type = 'number';
			input.min = '0';
			input.max = '255';
			input.step = '1';
			input.dataset.colorField = name.toLowerCase();
			return input;
		});
		const swatches = doc.createElement('div');
		swatches.className = 'more-colors-preview';
		const swatch = (text: string, name: string) => {
			const box = doc.createElement('span');
			box.dataset.colorPreview = name;
			const label = doc.createElement('span');
			label.textContent = text;
			swatches.append(label, box);
			return box;
		};
		this.#next = swatch('New', 'new');
		this.#current = swatch('Current', 'current');
		this.dialog.body.append(
			field('Hex', this.#hex),
			...CHANNELS.map((name, index) => field(name, this.#channels[index]!)),
			swatches,
		);
		this.#hex.addEventListener('input', () => this.#fromHex());
		for (const input of this.#channels) input.addEventListener('input', () => this.#fromChannels());
		// Typing stays in the fields; Enter accepts, as in Office's dialog.
		this.dialog.dialog.addEventListener('keydown', (event) => {
			if (event.key === 'Enter') {
				event.preventDefault();
				this.#accept();
			}
			if (event.key !== 'Escape' && event.key !== 'Tab') event.stopPropagation();
		});
	}
	/** Opens on `current` (`#rrggbb`; black when it is not a colour) and reports the choice. */
	open(current: string | undefined, pick: (color: string) => void): void {
		const start = normalizePaletteHex(current ?? '') ?? '#000000';
		this.#pick = pick;
		this.#hex.value = start.toUpperCase();
		this.#fromHex();
		this.#current.style.background = start;
		this.dialog.show();
		this.#hex.focus();
		this.#hex.select();
	}
	close(): void {
		this.#pick = undefined;
		this.dialog.close();
	}
	#value(): string | undefined {
		const text = this.#hex.value.trim();
		return /^#?[0-9a-f]{6}$/i.test(text) ? normalizePaletteHex(text) : undefined;
	}
	#fromHex(): void {
		const color = this.#value();
		if (color)
			this.#channels.forEach((input, index) => {
				input.value = String(parseInt(color.slice(1 + index * 2, 3 + index * 2), 16));
			});
		this.#preview(color);
	}
	#fromChannels(): void {
		const values = this.#channels.map((input) => Number(input.value));
		const valid = this.#channels.every(
			(input, index) =>
				input.value.trim() !== '' &&
				Number.isInteger(values[index]) &&
				values[index]! >= 0 &&
				values[index]! <= 255,
		);
		if (valid)
			this.#hex.value =
				`#${values.map((value) => value.toString(16).padStart(2, '0')).join('')}`.toUpperCase();
		this.#preview(valid ? this.#value() : undefined);
	}
	#preview(color: string | undefined): void {
		this.#next.style.background = color ?? 'transparent';
		this.dialog.error.textContent = '';
	}
	#accept(): void {
		const color = this.#value();
		const valid = this.#channels.every((input) => {
			const value = Number(input.value);
			return input.value.trim() !== '' && Number.isInteger(value) && value >= 0 && value <= 255;
		});
		if (!color || !valid) {
			this.dialog.error.textContent =
				'Enter the colour as #RRGGBB, or Red, Green and Blue from 0 to 255.';
			return;
		}
		const pick = this.#pick;
		this.close();
		pick?.(color);
	}
}
