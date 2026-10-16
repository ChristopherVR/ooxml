/**
 * The inspector's baseline for bare form markup.
 *
 * Many inspector sections (chart data, the advanced and exhaustive chart
 * blocks, text effects, accessibility, action settings, media, connector
 * arrows) build their fields as a bare `<label>caption<control></label>` with
 * no class, and some use a bare `<textarea>`, `<fieldset>` or native
 * `<input>`. Nothing in this stylesheet styled them, so their look came from
 * whatever the host page happened to ship: inside the vanilla demo Tailwind's
 * preflight stripped every input border (the fields were invisible), and in a
 * host without a reset, such as the OOXML Office suite, the browser defaults
 * showed through (a monospace textarea, grooved fieldset borders, white
 * inputs) while every label ran into its neighbours as one line of inline
 * text. A component must not depend on its host's reset, so the baseline lives
 * here.
 *
 * Every rule is one class plus `:where(...)`, so it is emitted FIRST (see
 * `buildViewerCss`) and any later class rule of the same or higher
 * specificity (`.pptxv-field-input`, the table-cell panel, the smartart
 * fields, ...) still wins. Being unlayered, it also beats a host reset that
 * sits in a cascade layer, as Tailwind 4's preflight does.
 */
export const INSPECTOR_BASE_CSS = `
/* ── Inspector baseline for bare form markup ─────────────────────────── */
.pptxv-inspector :where(input:not([type='checkbox']):not([type='radio']):not([type='color']):not([type='range']):not([type='file']):not([type='button']):not([type='submit']):not([type='reset']):not([type='hidden']), textarea) {
	box-sizing: border-box;
	min-width: 0;
	height: 26px;
	margin: 0;
	padding: 2px 6px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-background);
	color: var(--pptx-foreground);
	font: inherit;
	font-size: 12px;
}
.pptxv-inspector :where(textarea) {
	display: block;
	width: 100%;
	height: auto;
	min-height: 52px;
	resize: vertical;
}
.pptxv-inspector :where(input, textarea)::placeholder { color: var(--pptx-muted-foreground); opacity: 1; }
.pptxv-inspector :where(input, textarea):disabled { opacity: 0.5; }
.pptxv-inspector :where(input[type='color']) {
	box-sizing: border-box;
	flex: none;
	width: 28px;
	height: 20px;
	margin: 0;
	padding: 0;
	border: 1px solid var(--pptx-border);
	border-radius: 4px;
	background: none;
	cursor: pointer;
}
.pptxv-inspector :where(fieldset) {
	min-width: 0;
	margin: 0 0 8px;
	padding: 0;
	border: 0;
}
.pptxv-inspector :where(legend) {
	margin-bottom: 4px;
	padding: 0;
	color: var(--pptx-foreground);
	font-size: 11px;
	font-weight: 600;
}

/* One field per row: caption on the left, control on the right. A checkbox
   or radio goes first and its caption follows it, as in PowerPoint's panes.
   Only unclassed labels (and the chart helpers' field classes, which carry no
   rule of their own) take part: a classed label such as \`.pptxv-field\` has its
   own layout, and these flex declarations would otherwise leak into it. */
.pptxv-inspector-section :where(label:not([class]), .pptxv-exhaustive-field, .pptxv-chart-filters-range-field) {
	display: flex;
	flex-wrap: wrap;
	align-items: center;
	gap: 4px 8px;
	min-width: 0;
	margin-bottom: 6px;
	color: var(--pptx-muted-foreground);
	font-size: 11px;
}
/* A field toggled off with the hidden attribute must stay off: the flex
   display above would otherwise override the user-agent display: none. */
.pptxv-inspector-section :where(label:not([class]), .pptxv-exhaustive-field, .pptxv-chart-filters-range-field)[hidden] { display: none; }
.pptxv-inspector-section :where(label:not([class]), .pptxv-exhaustive-field, .pptxv-chart-filters-range-field) > :where(input, textarea, select, pptx-ui-select) {
	flex: 0 1 60%;
	margin-left: auto;
}
.pptxv-inspector-section :where(label:not([class]), .pptxv-exhaustive-field, .pptxv-chart-filters-range-field) > :where(input[type='color']) {
	flex: none;
	margin-left: auto;
}
.pptxv-inspector-section :where(label:not([class]), .pptxv-exhaustive-field, .pptxv-chart-filters-range-field) > :where(pptx-ui-checkbox, pptx-ui-radio, input[type='checkbox'], input[type='radio']) {
	flex: none;
	order: -1;
}
.pptxv-inspector-section :where(label:not([class]), .pptxv-exhaustive-field, .pptxv-chart-filters-range-field) > :where(textarea) {
	flex-basis: 100%;
	margin-left: 0;
}

/* Chart sub-blocks: a gap and a rule between the groups instead of one run. */
.pptxv-inspector :where(.pptxv-chart-advanced, .pptxv-chart-exhaustive, .pptxv-chart-subtype, .pptxv-chart-filters, .pptxv-chart-usershapes) {
	margin-top: 8px;
	padding-top: 8px;
	border-top: 1px solid var(--pptx-border);
}
.pptxv-inspector :where(.pptxv-chart-advanced, .pptxv-chart-exhaustive, .pptxv-chart-subtype, .pptxv-chart-filters, .pptxv-chart-usershapes)[hidden] { display: none; }
.pptxv-inspector :where(.pptxv-chart-advanced, .pptxv-chart-exhaustive, .pptxv-chart-filters, .pptxv-chart-usershapes) :where(button:not([class]), .pptxv-chart-clear-series-color) {
	min-height: 24px;
	padding: 2px 8px;
	border: 1px solid var(--pptx-border);
	border-radius: var(--pptx-radius);
	background: var(--pptx-secondary);
	color: var(--pptx-secondary-foreground);
	font: inherit;
	font-size: 11px;
	cursor: pointer;
}
.pptxv-inspector :where(.pptxv-chart-clear-series-color) { margin-bottom: 6px; }
.pptxv-inspector :where(.pptxv-chart-clear-series-color)[hidden] { display: none; }
`;
