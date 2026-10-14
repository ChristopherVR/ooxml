/** Fixed-size page shells keep scroll geometry stable while only nearby content is mounted. */
export class PrintPageWindow {
	private pages: { sheet: HTMLElement; draw: () => () => void }[] = [];
	private mounted = new Map<number, () => void>();
	private revision = 0;

	constructor(
		private readonly root: HTMLElement,
		private readonly virtual: boolean,
	) {}

	add(sheet: HTMLElement, draw: () => () => void): void {
		this.pages.push({ sheet, draw });
		this.root.append(sheet);
	}

	mount(index: number): void {
		if (this.mounted.has(index)) return;
		const page = this.pages[index];
		if (!page) return;
		this.mounted.set(index, page.draw());
		this.changed();
	}

	private unmount(index: number): void {
		this.mounted.get(index)?.();
		this.mounted.delete(index);
		this.pages[index]?.sheet.replaceChildren();
		this.changed();
	}

	private changed(): void {
		this.root.dataset.renderVersion = String(++this.revision);
	}

	show(center: number, pinned?: number): void {
		if (!this.virtual) return this.all();
		const first = Math.max(0, center - 2);
		const last = Math.min(this.pages.length - 1, center + 2);
		for (const index of this.mounted.keys())
			if ((index < first || index > last) && index !== pinned) this.unmount(index);
		for (let index = first; index <= last; index++) this.mount(index);
		if (pinned !== undefined) this.mount(pinned);
	}

	all(): void {
		for (let index = 0; index < this.pages.length; index++) this.mount(index);
	}

	snapshot(index: number): HTMLElement | undefined {
		const existed = this.mounted.has(index);
		this.mount(index);
		const clone = this.pages[index]?.sheet.cloneNode(true) as HTMLElement | undefined;
		if (!existed) this.unmount(index);
		return clone;
	}
}
