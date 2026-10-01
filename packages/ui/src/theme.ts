/**
 * Theme-token foundation. Custom properties cross shadow boundaries, so one `<style>` in the
 * document head themes every control. Every token also has a fallback inside each control, so
 * installing the theme is optional; a host overrides any token on `:root` or on a container.
 */
export const THEME_STYLE_ID = 'office-ui-theme';

const DARK = `
	--office-foreground: #f9fafb;
	--office-muted-foreground: #9ca3af;
	--office-background: #111827;
	--office-surface: #1f2937;
	--office-border: #374151;
	--office-accent: #6366f1;
	--office-ring: #818cf8;
	--office-danger: #f87171;`;

export const THEME_CSS = `
:root {
	--office-foreground: #1f2937;
	--office-muted-foreground: #6b7280;
	--office-background: #ffffff;
	--office-surface: #f3f4f6;
	--office-border: #d1d5db;
	--office-accent: #2563eb;
	--office-accent-foreground: #ffffff;
	--office-ring: #2563eb;
	--office-danger: #b91c1c;
	--office-target-size: 28px;
	--office-radius: 4px;
	--office-font: system-ui, -apple-system, "Segoe UI", sans-serif;
}
@media (prefers-color-scheme: dark) {
	:root:not([data-office-theme="light"]) {${DARK}
	}
}
:root[data-office-theme="dark"] {${DARK}
}
@media (pointer: coarse), (max-width: 767px) {
	:root { --office-target-size: 44px; }
}
@media (forced-colors: active) {
	:root {
		--office-foreground: CanvasText;
		--office-muted-foreground: GrayText;
		--office-background: Canvas;
		--office-surface: Canvas;
		--office-border: CanvasText;
		--office-accent: Highlight;
		--office-accent-foreground: HighlightText;
		--office-ring: Highlight;
		--office-danger: CanvasText;
	}
}
`;

/** Idempotent. Browser-only; returns false (and does nothing) without a document. */
export function installOfficeUiTheme(doc: Document | undefined = globalThis.document): boolean {
	if (!doc?.head) return false;
	if (!doc.getElementById(THEME_STYLE_ID)) {
		const style = doc.createElement('style');
		style.id = THEME_STYLE_ID;
		style.textContent = THEME_CSS;
		doc.head.append(style);
	}
	return true;
}
