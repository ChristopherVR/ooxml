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
	applyTeamsProps,
	defineTeamsApp,
	listenTeamsEvents,
	type FileOpeners,
	type FileEmbeds,
	type FileUploader,
	type OpenFileDetail,
	type TeamsApp,
	type TeamsProps,
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
	@Output() ready = new EventEmitter<{ user: { id: string; name: string } }>();
	@Output() openFile = new EventEmitter<{
		detail: OpenFileDetail;
		event: CustomEvent<OpenFileDetail>;
	}>();
	@Output() configChange = new EventEmitter<{ config: TeamsServerConfig }>();
	@ViewChild('host', { static: true }) host!: ElementRef<TeamsApp>;

	private applied: TeamsProps = {};
	private stop = (): void => {};

	private current(): TeamsProps {
		return {
			...(this.workspaceId !== undefined ? { workspaceId: this.workspaceId } : {}),
			...(this.userName !== undefined ? { userName: this.userName } : {}),
			...(this.userId !== undefined ? { userId: this.userId } : {}),
			...(this.config !== undefined ? { config: this.config } : {}),
			...(this.uploadFile ? { uploadFile: this.uploadFile } : {}),
			...(this.openers ? { openers: this.openers } : {}),
			...(this.embeds ? { embeds: this.embeds } : {}),
			onReady: (d) => this.ready.emit(d),
			onOpenFile: (detail, event) => this.openFile.emit({ detail, event }),
			onConfigChange: (d) => this.configChange.emit(d),
		};
	}
	private sync(): void {
		const next = this.current();
		applyTeamsProps(this.host.nativeElement, next, this.applied);
		this.applied = next;
	}
	ngAfterViewInit(): void {
		defineTeamsApp();
		this.stop = listenTeamsEvents(this.host.nativeElement, () => this.current());
		this.sync();
	}
	ngOnChanges(): void {
		if (this.host) this.sync();
	}
	ngOnDestroy(): void {
		this.stop();
	}
}
export * from './service';
