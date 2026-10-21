import * as officeTheme from 'ooxml-ui/theme';
export type BuildingBlockKind = 'toolbar' | 'formula-bar' | 'tabs';
/** Lazy lookup keeps existing viewer imports compatible with an older Office UI runtime. */
export function prepareBlock(value: unknown, kind: BuildingBlockKind): void {
	if (!value || typeof value !== 'object' || !('ownerDocument' in value)) return;
	const host = value as HTMLElement;
	const prepare = Reflect.get(officeTheme, 'prepareOfficeUiBlock') as
		| ((host: HTMLElement, kind: BuildingBlockKind, product: 'xlsx') => void)
		| undefined;
	if (typeof prepare === 'function') prepare(host, kind, 'xlsx');
	else {
		officeTheme.installOfficeUiTheme(host.ownerDocument);
		host.setAttribute('data-office-block', kind);
		host.setAttribute('data-office-product', 'xlsx');
	}
}
