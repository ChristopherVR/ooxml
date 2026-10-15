/**
 * settings-dialog.component.ts: the File > Options dialog (Angular port of
 * React's `SettingsDialog.tsx`).
 *
 * The shared `office-ui-options-dialog`, driven by `bindPptxOptionsDialog`: the ten categories,
 * their settings, live apply, Cancel and Reset all live in `ooxml-ui/pptx`. This component only
 * slots the panes PowerPoint has on top of the generic rows (themes, fonts, languages, Customize
 * Ribbon, Quick Access Toolbar, Add-ins, AI) while their category is open.
 */
import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	DestroyRef,
	effect,
	inject,
	input,
	output,
	signal,
	untracked,
	viewChild,
} from '@angular/core';
import type { ElementRef } from '@angular/core';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import { bindPptxOptionsDialog, THEME_CATALOG } from 'ooxml-ui/pptx';
import type {
	PptxOptionsDialogBinding,
	PptxOptionsDialogElement,
	PptxOptionsDialogInput,
	ThemeCatalogEntry,
	ViewerAddinStatus,
	ViewerOptionPrimitive,
	ViewerOptions,
	ViewerOptionsGroupId,
} from 'ooxml-ui/pptx';
import { LOCALE_CATALOG } from 'ooxml-ui/pptx/i18n';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';

import { AiSettingsSectionComponent } from './ai/ai-settings-section.component';
import { OptionsAddInsPaneComponent } from './options-add-ins-pane.component';
import { OptionsQuickAccessPaneComponent } from './options-quick-access-pane.component';
import { OptionsRibbonPaneComponent } from './options-ribbon-pane.component';
import type { RibbonTabHiddenChange } from './options-ribbon-pane.component';
import { SettingsAppearanceTabComponent } from './settings-appearance-tab.component';
import { SettingsCustomFontsComponent } from './settings-custom-fonts.component';
import { SettingsLanguageTabComponent } from './settings-language-tab.component';
import { translationsSignal } from './translations-signal';
import { injectResolvedCustomization } from './viewer-customization.service';

export type { ViewerSettings } from 'ooxml-ui/pptx';

/** One dialog-level option edit (group + key + new primitive value). */
export interface OptionValueChange {
	group: ViewerOptionsGroupId;
	key: string;
	value: ViewerOptionPrimitive;
}

@Component({
	selector: 'pptx-settings-dialog',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	imports: [
		TranslatePipe,
		OptionsRibbonPaneComponent,
		OptionsQuickAccessPaneComponent,
		OptionsAddInsPaneComponent,
		SettingsAppearanceTabComponent,
		SettingsCustomFontsComponent,
		SettingsLanguageTabComponent,
		AiSettingsSectionComponent,
	],
	template: `
		@if (open()) {
			<pptx-ui-options-dialog #dialog>
				@switch (category()) {
					@case ('general') {
						<div slot="block-themePicker">
							<pptx-settings-appearance-tab
								[themes]="availableThemes()"
								[activeKey]="themeKey()"
								(select)="themeKeySelect.emit($event)"
							/>
						</div>
						<div slot="block-customFonts">
							<pptx-settings-custom-fonts
								[enabled]="options().general.enableCustomFontUpload"
								[families]="customFontFamilies()"
								(registered)="customFontRegistered.emit($event)"
							/>
						</div>
					}
					@case ('save') {
						<div slot="block-clearCache" class="pptx-ng-options-clear-cache">
							<p>{{ 'pptx.options.save.clearCacheDescription' | translate }}</p>
							<button type="button" (click)="clearCache.emit()">
								{{ 'pptx.options.save.clearCacheNow' | translate }}
							</button>
						</div>
					}
					@case ('language') {
						<div slot="block-language">
							<pptx-settings-language-tab
								[locales]="availableLocales()"
								[activeCode]="localeCode()"
								(select)="localeSelect.emit($event)"
							/>
						</div>
					}
					@case ('ribbon') {
						<div slot="pane-ribbon">
							<pptx-options-ribbon-pane
								[options]="options()"
								(tabHiddenChange)="ribbonTabHiddenChange.emit($event)"
								(resetRibbon)="resetOptions.emit('ribbon')"
							/>
						</div>
					}
					@case ('quickAccess') {
						<div slot="pane-quickAccess">
							<pptx-options-quick-access-pane
								[options]="options()"
								(commandsChange)="quickAccessCommandsChange.emit($event)"
							/>
						</div>
					}
					@case ('addIns') {
						<div slot="pane-addIns">
							<pptx-options-add-ins-pane [addinStatus]="addinStatus()" />
						</div>
					}
					@case ('ai') {
						<div slot="pane-ai">
							<pptx-ai-settings-section />
						</div>
					}
				}
			</pptx-ui-options-dialog>
		}
	`,
	styles: [
		`
			.pptx-ng-options-clear-cache p {
				margin: 0 0 8px;
				color: var(--pptx-muted-foreground);
				font-size: 11px;
			}
			.pptx-ng-options-clear-cache button {
				padding: 6px 14px;
				border: 1px solid var(--pptx-border);
				border-radius: 6px;
				background: transparent;
				color: var(--pptx-foreground);
				font-size: 12px;
				cursor: pointer;
			}
		`,
	],
})
export class SettingsDialogComponent {
	readonly open = input(false);
	/** Full File > Options snapshot rendered by every pane. */
	readonly options = input.required<ViewerOptions>();
	/** Selected theme catalog key, for General > Appearance. */
	readonly themeKey = input<string>('default');
	readonly availableThemes = input<readonly ThemeCatalogEntry[]>(THEME_CATALOG);
	/** Active locale code, for the Language category. */
	readonly localeCode = input<string>('en');
	readonly availableLocales = input<readonly LocaleCatalogEntry[]>(LOCALE_CATALOG);
	/** Availability flags for the Add-ins pane (unset ids default to active). */
	readonly addinStatus = input<ViewerAddinStatus | undefined>(undefined);
	/**
	 * Whether to show the bespoke "AI" category (chat-log export), gated by the
	 * host on the `ai` config. Absent by default (no AI category).
	 */
	readonly aiExportVisible = input<boolean>(false);

	readonly optionChange = output<OptionValueChange>();
	/** Restore a snapshot wholesale (Cancel semantics). */
	readonly restoreOptions = output<ViewerOptions>();
	readonly ribbonTabHiddenChange = output<RibbonTabHiddenChange>();
	readonly quickAccessCommandsChange = output<string[]>();
	/** Reset one tab-group (or everything when `undefined`). */
	readonly resetOptions = output<ViewerOptionsGroupId | undefined>();
	readonly clearCache = output<void>();
	readonly themeKeySelect = output<string>();
	/** Families registered this session via the Fonts section. */
	readonly customFontFamilies = input<readonly string[]>([]);
	/** A font file was registered; the ribbon adds the family to its list. */
	readonly customFontRegistered = output<string>();
	readonly localeSelect = output<string>();
	readonly close = output<void>();

	private readonly translate = inject(TranslateService);
	private readonly translations = translationsSignal(this.translate);
	private readonly customization = injectResolvedCustomization();
	private readonly dialog = viewChild<ElementRef<PptxOptionsDialogElement>>('dialog');
	private binding: PptxOptionsDialogBinding | null = null;
	private readonly t = (key: string): string => this.translate.instant(key);
	/** The open category, for the panes slotted into the shared dialog. */
	protected readonly category = signal('general');

	constructor() {
		// The element exists only while the dialog is open: bind it when it appears.
		effect(() => {
			const element = this.dialog()?.nativeElement;
			untracked(() => {
				this.unbind();
				if (element) this.bind(element);
			});
		});
		// Follow every option, language, customisation and AI change.
		effect(() => {
			this.dialog();
			const input = this.dialogInput();
			untracked(() => this.binding?.sync(input));
		});
		inject(DestroyRef).onDestroy(() => this.unbind());
	}

	/** What the shared dialog renders from, read off this component's inputs. */
	private dialogInput(): PptxOptionsDialogInput {
		return {
			options: this.options(),
			t: this.t,
			languageKey: String(this.translations()),
			customization: this.customization(),
			aiEnabled: this.aiExportVisible(),
		};
	}

	private bind(element: PptxOptionsDialogElement): void {
		this.binding = bindPptxOptionsDialog(element, {
			setOption: (group, key, value) => this.optionChange.emit({ group, key, value }),
			restoreOptions: (snapshot) => this.restoreOptions.emit(snapshot),
			resetOptions: () => this.resetOptions.emit(undefined),
			close: () => this.close.emit(),
			categoryChanged: (id) => this.category.set(id),
		});
		this.binding.sync(untracked(() => this.dialogInput()));
	}

	private unbind(): void {
		this.binding?.dispose();
		this.binding = null;
	}
}
