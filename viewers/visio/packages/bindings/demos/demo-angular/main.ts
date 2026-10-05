import '@angular/compiler';
import { ChangeDetectorRef, Component, ViewChild, inject, type AfterViewInit } from '@angular/core';
import { bootstrapApplication } from '@angular/platform-browser';
import { VisioViewerComponent } from '../../src/angular.js';
import { createWorkspace } from '../../../../demo/workspace.js';

// The Angular demo: <visio-viewer-host> with inputs, a @ViewChild handle and change detection.
const workspace = createWorkspace();

@Component({
	selector: 'visio-demo-app',
	standalone: true,
	imports: [VisioViewerComponent],
	template: '<visio-viewer-host [document]="document" [events]="events"></visio-viewer-host>',
})
class AppComponent implements AfterViewInit {
	@ViewChild(VisioViewerComponent) viewer?: VisioViewerComponent;
	document = workspace.initialDocument;
	readonly events = workspace.events;
	private readonly changes = inject(ChangeDetectorRef);
	ngAfterViewInit(): void {
		if (!this.viewer) return;
		workspace.attach(this.viewer, (next) => {
			this.document = next;
			this.changes.detectChanges();
		});
	}
}

const host = window.document.getElementById('viewer')!;
host.append(window.document.createElement('visio-demo-app'));
void bootstrapApplication(AppComponent);
