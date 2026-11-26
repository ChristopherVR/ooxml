import { bootstrapApplication } from '@angular/platform-browser';

import '../../shared/zod-jitless';
import 'zone.js';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';

import { enableDemoAddInPreferences } from '../../shared/demo-add-ins';
// Test-only: `?elementRenderFault=<id>` makes that element's renderer throw (e2e).
import '../../shared/element-render-fault';
import { externalSessionRequested } from '../../shared/host-owned-collaboration';
import { sameBrowserRoom } from '../../shared/same-browser-collaboration';
import { AppComponent } from './app.component';
import { BuildingBlocksDemoComponent } from './building-blocks-demo.component';
import { HostOwnedDemoComponent } from './host-owned-demo.component';
import { i18nProviders } from './i18n';
import { SameBrowserDemoComponent } from './same-browser-demo.component';

import '../../shared/dropzone.css';

if (!completeDemoSignInPopup()) {
	enableDemoAddInPreferences(document.body);
	bootstrapApplication(
		new URLSearchParams(location.search).get('blocks') === '1'
			? BuildingBlocksDemoComponent
			: externalSessionRequested()
				? HostOwnedDemoComponent
				: sameBrowserRoom()
					? SameBrowserDemoComponent
					: AppComponent,
		{ providers: [i18nProviders] },
	).catch((err) => console.error(err));
}
