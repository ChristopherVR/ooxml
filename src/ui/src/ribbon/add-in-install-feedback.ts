import type { LocalAddIn } from 'ooxml-core/automation/add-ins';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import type { AddInDialogHost } from './add-in-dialog';
import { MANIFEST_TAB_PREFIX } from './add-in-manifest-tabs';

export function addInInstallFeedback(
	host: AddInDialogHost,
	addIn: LocalAddIn,
	t: OfficeServiceText,
): boolean {
	const { manifest } = addIn;
	let status = host.querySelector<HTMLElement>('[data-office-add-in-status]');
	if (!status) {
		status = host.ownerDocument.createElement('div');
		status.dataset.officeAddInStatus = '';
		status.setAttribute('role', 'status');
		status.style.cssText =
			'position:absolute;width:1px;height:1px;overflow:hidden;clip-path:inset(50%)';
		host.append(status);
	}
	status.textContent = '';
	const region = status;
	host.ownerDocument.defaultView?.requestAnimationFrame(() => {
		region.textContent = t('{name} installed.').replace('{name}', manifest.name);
	});
	const ribbon = host.querySelector('office-ui-ribbon, pptx-ui-ribbon-tabs');
	const prefix = `${MANIFEST_TAB_PREFIX}${manifest.id}.`;
	const tab = [
		...(ribbon?.shadowRoot?.querySelectorAll<HTMLButtonElement>('[role="tab"]') ?? []),
	].find(({ dataset }) => dataset.tab?.startsWith(prefix));
	if (!tab) return false;
	tab.click();
	tab.focus();
	return true;
}
