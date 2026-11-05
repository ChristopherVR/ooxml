import {
	ChangeDetectionStrategy,
	Component,
	DestroyRef,
	ElementRef,
	effect,
	inject,
	input,
	output,
	untracked,
} from '@angular/core';
import { toSignal } from '@angular/core/rxjs-interop';
import { TranslateService } from '@ngx-translate/core';
import { mountCloudFilesPlace } from 'ooxml-ui/pptx';
import type {
	CloudFileProvider,
	CloudFilesPlace,
	CloudFilesPlaceOptions,
	CloudOpenedFile,
} from 'ooxml-ui/pptx';
import { merge, map, of, startWith } from 'rxjs';

/**
 * One File > Open place for a cloud provider: the shared `pptx-ui-cloud-files` browser, mounted
 * into this element (`data-pptx-cloud-place`) by the shared `mountCloudFilesPlace`.
 */
@Component({
	selector: 'pptx-cloud-files-place',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: {
		'[attr.data-pptx-cloud-place]': 'provider().id',
		style: 'display: block; margin-top: 24px; max-width: 900px',
	},
	template: '',
})
export class CloudFilesPlaceComponent {
	readonly provider = input.required<CloudFileProvider>();
	/** The user opened a file in this place. */
	readonly open = output<CloudOpenedFile>();

	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	/** Optional, like the backstage: without ngx-translate the element keeps its English labels. */
	private readonly translate = inject(TranslateService, { optional: true });
	/** Changes on language/dictionary updates so the element's labels follow the viewer. */
	private readonly translations = toSignal(
		this.translate
			? merge(this.translate.onLangChange, this.translate.onTranslationChange).pipe(
					map(() => Date.now()),
					startWith(0),
				)
			: of(0),
		{ initialValue: 0 },
	);
	private place: CloudFilesPlace | null = null;
	private mountedProvider: CloudFileProvider | null = null;

	constructor() {
		effect(() => {
			const provider = this.provider();
			this.translations();
			const locale = this.translate?.currentLang() ?? undefined;
			untracked(() => this.apply(this.options(provider, locale)));
		});
		inject(DestroyRef).onDestroy(() => this.unmount());
	}

	/** Mount the browser for `options.provider`, or update the one already showing it. */
	apply(options: CloudFilesPlaceOptions): void {
		if (this.place && this.mountedProvider === options.provider) {
			this.place.update(options);
			return;
		}
		this.unmount();
		this.place = mountCloudFilesPlace(this.host.nativeElement, options);
		this.mountedProvider = options.provider;
	}

	private options(provider: CloudFileProvider, locale: string | undefined): CloudFilesPlaceOptions {
		const translate = this.translate;
		return {
			provider,
			locale,
			translate: translate
				? (key, params) => translate.instant(key, params) as string
				: (key) => key,
			onOpen: (file) => this.open.emit(file),
		};
	}

	private unmount(): void {
		this.place?.destroy();
		this.place = null;
		this.mountedProvider = null;
	}
}
