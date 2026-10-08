import {
	CUSTOM_ELEMENTS_SCHEMA,
	Component,
	ElementRef,
	EventEmitter,
	Input,
	Output,
	ViewChild,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
} from '@angular/core';
import type { TeamsServerConfig } from 'ooxml-core/teams';
import {
	bindTeams,
	defineTeamsApp,
	pickTeamsProps,
	type FileOpeners,
	type FileEmbeds,
	type FileUploader,
	type OpenFileDetail,
	type TeamsApp,
	type TeamsBinding,
	type TeamsElementProps,
} from 'teams-viewer';

/** `<teams-workspace workspaceId="acme" userName="Ada" [config]="cfg" (openFile)="..." />` */
@Component({
	selector: 'teams-workspace',
	standalone: true,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `<teams-app #host style="display:block;height:100%"></teams-app>`,
	styles: [':host { display: block; height: 100%; }'],
})
export class TeamsWorkspaceComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input() workspaceId?: string;
	@Input() userName?: string;
	@Input() userId?: string;
	@Input() config?: TeamsServerConfig | null;
	@Input() uploadFile?: FileUploader;
	@Input() openers?: FileOpeners;
	@Input() embeds?: FileEmbeds;
	/** Classes for the inner <teams-app> (a plain `class` attribute styles the host element). */
	@Input() className?: string;
	@Output() ready = new EventEmitter<{ user: { id: string; name: string } }>();
	/**
	 * Angular outputs carry one value, so this emits `{ detail, event }` where the other bindings
	 * call `onOpenFile(detail, event)`. Call `event.preventDefault()` to take over opening a file.
	 */
	@Output() openFile = new EventEmitter<{
		detail: OpenFileDetail;
		event: CustomEvent<OpenFileDetail>;
	}>();
	@Output() configChange = new EventEmitter<{ config: TeamsServerConfig }>();
	@ViewChild('host', { static: true }) host!: ElementRef<TeamsApp>;

	private binding: TeamsBinding | undefined;

	private current(): TeamsElementProps {
		return {
			...pickTeamsProps(this),
			onReady: (d) => this.ready.emit(d),
			onOpenFile: (detail, event) => this.openFile.emit({ detail, event }),
			onConfigChange: (d) => this.configChange.emit(d),
		};
	}
	private sync(): void {
		this.binding?.update(this.current());
	}
	ngAfterViewInit(): void {
		defineTeamsApp();
		this.binding = bindTeams(this.host.nativeElement, () => this.current());
		this.sync();
	}
	ngOnChanges(): void {
		if (this.host) this.sync();
	}
	ngOnDestroy(): void {
		this.binding?.destroy();
	}
}
export * from './service';
