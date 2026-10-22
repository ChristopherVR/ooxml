import 'zone.js';
// Load the Angular JIT compiler before bootstrapping.
//
// This demo runs on Vite 8 with @analogjs/vite-plugin-angular, whose ahead-of-
// time (AOT) transform does not run against local components in this toolchain
// combination — esbuild transpiles `@Component` with standard decorators, so at
// runtime Angular has no compiled component definition and falls back to JIT.
// Pulling in '@angular/compiler' makes that fallback work (the exact remedy the
// runtime error recommends). JIT is fine for a local demo; the published
// `pptx-angular-viewer` library itself is AOT-compiled by ng-packagr and never
// depends on this.
import '@angular/compiler';
import { bootstrapApplication } from '@angular/platform-browser';

import { externalSessionRequested } from '../../shared/host-owned-collaboration';
import { sameBrowserRoom } from '../../shared/same-browser-collaboration';
import { AppComponent } from './app.component';
import { HostOwnedDemoComponent } from './host-owned-demo.component';
import { i18nProviders } from './i18n';
import { SameBrowserDemoComponent } from './same-browser-demo.component';

import '../../shared/dropzone.css';

bootstrapApplication(
	externalSessionRequested()
		? HostOwnedDemoComponent
		: sameBrowserRoom()
			? SameBrowserDemoComponent
			: AppComponent,
	{ providers: [i18nProviders] },
).catch((err) => console.error(err));
