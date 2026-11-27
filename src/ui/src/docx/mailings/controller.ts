import type { TextRun } from 'ooxml-core/docx';
import {
	checkMailMerge,
	mailMergeWithinWorkLimit,
	previewMailMerge,
	type MailMergeData,
	type MailMergeSettings,
} from 'ooxml-core/docx/mail-merge';
import { Fragment, Slice } from 'prosemirror-model';
import { closeHistory } from 'prosemirror-history';
import type { EditorCore } from '../editor-core';
import { focusView } from '../focus-view';
import type { LocalizationKey } from '../localization';
import {
	mailingsReason,
	syncMailingsRibbon,
	type MailingsCommand,
	type MailingsState,
} from '../mailings-ribbon';
import { runToInlineNodes } from '../run-adapter';
import { createMailingsDialog, dialogMessage, type MailingsDialog } from './dialog';
import { nativeFieldDialog, fieldDialog, matchFieldsDialog, rulesDialog } from './field-dialog';
import { MailingsProjection } from './projection';
import {
	allRecipients,
	selectedRecipients,
	snapshotRecipients,
	type RecipientSelection,
} from './recipient-data';
import { editRecipientList } from './recipient-dialog';
import {
	checkDialog,
	findDialog,
	finishDialog,
	sourceDialog,
	startDialog,
	type MergeType,
} from './workflow-dialogs';
import type { MailingsLabel } from './strings';
import { readLocalRecipients } from './recipient-file';
import { emailUnavailable } from './email';
import { finishWorkflow } from './finish';
import {
	applyDocument,
	emptyBody,
	envelopeDialog,
	labelsDialog,
	labelUpdateIssue,
	openLabels,
	startLabels,
	withStationery,
} from './stationery';

export class MailingsController {
	private source: MailMergeData | undefined;
	private selection = allRecipients();
	private recipients: MailMergeData = { records: [], fields: [] };
	private record = 1;
	private preview = false;
	private highlight = false;
	private busy = false;
	private epoch = 0;
	private dialog: MailingsDialog | undefined;
	private readonly projection: MailingsProjection;

	constructor(private readonly core: EditorCore) {
		this.projection = new MailingsProjection(core);
	}

	get state(): MailingsState {
		return {
			count: this.recipients.records.length,
			record: this.record,
			hasFields: Boolean(this.source?.fields?.length),
			hasSource: Boolean(this.source),
			preview: this.preview,
			highlight: this.highlight,
			busy: this.busy,
		};
	}

	disconnect(): void {
		this.epoch++;
		this.busy = false;
		this.dialog?.close();
		this.dialog = undefined;
		this.stopPreview();
	}

	stopPreview(): void {
		if (!this.preview) return;
		this.preview = false;
		this.projection.stop();
		syncMailingsRibbon(this.core);
	}

	run(key: MailingsCommand): void {
		if (mailingsReason(this.core, key)) return;
		const { core } = this;
		const { locale } = core;
		const fields = this.source?.fields ?? [];
		const mappings = core.model.mailMerge?.fieldMappings ?? [];
		const insert = (runs: TextRun[]) => {
			if (mailingsReason(core, key) || !runs.length) return;
			const view = core.view;
			if (!view?.editable) return;
			const { $from, $to } = view.state.selection;
			if (!$from.sameParent($to) || !$from.parent.inlineContent) return;
			const fragment = Fragment.fromArray(runs.flatMap((run) => runToInlineNodes(run)));
			view.dispatch(
				closeHistory(view.state.tr)
					.replaceSelection(new Slice(fragment, 0, 0))
					.scrollIntoView(),
			);
			focusView(view);
		};
		switch (key) {
			case 'start':
				this.show('Start Mail Merge', (dialog) =>
					startDialog(
						dialog,
						locale,
						core.model.mailMerge?.mainDocumentType,
						emptyBody(core),
						(type) => this.start(type),
					),
				);
				break;
			case 'envelopes':
				this.openEnvelope(false);
				break;
			case 'labels':
				this.openLabelSheet(false);
				break;
			case 'updateLabels':
				void withStationery(core, key, ({ updateLabels }) => {
					const next = updateLabels(core.model);
					if (next === core.model) this.show('Update Labels', labelUpdateIssue);
					else applyDocument(core, next);
				});
				break;
			case 'source':
				this.show('Select Recipients', (dialog) =>
					sourceDialog(dialog, core.mailMergeDataSource, (file) => {
						if (mailingsReason(core, key)) return;
						const provider = core.mailMergeDataSource;
						const read = file
							? () => readLocalRecipients(file)
							: provider && (async () => snapshotRecipients(await provider.read()));
						if (!read) return;
						void this.operation(
							'Select Recipients',
							'The recipient source could not be read. Use at most 1,000 rows, 64 columns, and a 5 MiB file.',
							async (current) => {
								const data = await read();
								if (!current()) return;
								this.source = data;
								this.select(allRecipients());
							},
						);
					}),
				);
				break;
			case 'recipients': {
				const source = this.source;
				if (source)
					this.show('Edit Recipient List', (dialog) =>
						editRecipientList(dialog, locale, source, this.selection, (settings) => {
							if (!mailingsReason(core, key)) this.select(settings);
						}),
					);
				break;
			}
			case 'highlight':
				this.highlight = !this.highlight;
				core.element.toggleAttribute('data-highlight-merge-fields', this.highlight);
				break;
			case 'field':
				this.show('Insert Merge Field', (dialog) => fieldDialog(dialog, locale, fields, insert));
				break;
			case 'match':
				this.show('Match Fields', (dialog) =>
					matchFieldsDialog(dialog, locale, fields, mappings, (fieldMappings) => {
						if (!mailingsReason(core, key))
							this.applySettings({ ...core.model.mailMerge, fieldMappings });
					}),
				);
				break;
			case 'address':
			case 'greeting':
				this.show(key === 'address' ? 'Address Block' : 'Greeting Line', (dialog) =>
					nativeFieldDialog(dialog, locale, key, fields, mappings, insert),
				);
				break;
			case 'rules':
				this.show('Rules', (dialog) => rulesDialog(dialog, locale, fields, insert));
				break;
			case 'preview':
				if (this.preview) this.stopPreview();
				else this.showRecord(this.record);
				break;
			case 'first':
				this.showRecord(1);
				break;
			case 'previous':
				this.showRecord(this.record - 1);
				break;
			case 'next':
				this.showRecord(this.record + 1);
				break;
			case 'last':
				this.showRecord(this.state.count);
				break;
			case 'find':
				this.show('Find Recipient', (dialog) =>
					findDialog(dialog, locale, this.recipients, this.preview ? this.record : 0, (record) => {
						if (!mailingsReason(core, key)) this.showRecord(record);
					}),
				);
				break;
			case 'check':
				if (this.withinLimit('Check for Errors'))
					this.show('Check for Errors', (dialog) =>
						checkDialog(dialog, checkMailMerge(core.model, this.recipients, { locale })),
					);
				break;
			case 'finish':
				this.show('Finish & Merge', (dialog) =>
					finishDialog(
						dialog,
						(destination) => {
							if (mailingsReason(core, key) || !this.withinLimit('Finish & Merge')) return;
							void this.operation(
								'Finish & Merge',
								'The merge could not be completed.',
								(current) =>
									finishWorkflow(
										core,
										this.recipients,
										destination,
										(title, build) => this.show(title, build),
										current,
										(model) => this.projection.print(model),
									),
							);
						},
						emailUnavailable(core, this.recipients),
					),
				);
		}
		syncMailingsRibbon(core);
	}

	private openEnvelope(merge: boolean): void {
		const { core } = this;
		const key = merge ? 'start' : 'envelopes';
		void withStationery(core, key, ({ addEnvelope, envelopeSizes }) => {
			this.show('Envelopes', (dialog) =>
				envelopeDialog(dialog, core.locale, envelopeSizes, (options) => {
					if (mailingsReason(core, key) || (merge && !emptyBody(core))) return;
					applyDocument(core, addEnvelope(core.model, options, merge ? 'merge' : 'prepend'), merge);
				}),
			);
		});
	}

	private openLabelSheet(merge: boolean): void {
		const { core } = this;
		const key = merge ? 'start' : 'labels';
		void withStationery(core, key, ({ createLabels, labelVendors, searchLabels }) => {
			this.show('Labels', (dialog) =>
				labelsDialog(dialog, { labelVendors, searchLabels }, merge, (options) => {
					if (mailingsReason(core, key) || (merge && !emptyBody(core))) return;
					const sheet = createLabels(options);
					if (merge) startLabels(core, sheet);
					else
						void this.operation('Labels', 'The label document could not be created.', (current) =>
							openLabels(core, sheet, current),
						);
				}),
			);
		});
	}

	private start(type: MergeType | undefined): void {
		const { core } = this;
		if (mailingsReason(core, 'start') || (type && type === core.model.mailMerge?.mainDocumentType))
			return;
		if (!type) {
			this.applySettings(undefined);
			this.source = undefined;
			this.select(allRecipients());
		} else if (type === 'envelopes' && emptyBody(core)) {
			this.openEnvelope(true);
		} else if (type === 'mailingLabels' && emptyBody(core)) {
			this.openLabelSheet(true);
		} else if (type === 'formLetters' || type === 'catalog') {
			this.applySettings({ ...core.model.mailMerge, mainDocumentType: type });
		}
	}

	private select(settings: RecipientSelection): void {
		this.stopPreview();
		this.selection = settings;
		this.recipients = this.source
			? selectedRecipients(this.source, settings, this.core.locale)
			: { records: [], fields: [] };
		this.record = 1;
		syncMailingsRibbon(this.core);
	}

	private applySettings(mailMerge: MailMergeSettings | undefined): void {
		const view = this.core.view;
		if (view?.editable)
			view.dispatch(closeHistory(view.state.tr).setDocAttribute('mailMerge', mailMerge ?? null));
	}

	private showRecord(record: number): void {
		if (!this.withinLimit('Preview Results')) return;
		const { core } = this;
		const result = previewMailMerge(core.model, this.recipients, record, { locale: core.locale });
		if (!this.preview) core.parts.closeOpenStory(false);
		this.projection.show(result.document, record, this.state.count, result.skipped);
		this.record = record;
		this.preview = true;
		syncMailingsRibbon(core);
		if (!core.shell.toolbar?.contains(core.element.shadowRoot?.activeElement ?? null))
			core.shell.toolbar?.querySelector<HTMLElement>('[data-mailings="preview"]')?.focus();
	}

	private withinLimit(title: LocalizationKey): boolean {
		if (mailMergeWithinWorkLimit(this.core.model, this.recipients)) return true;
		this.message(
			title,
			'This merge exceeds the editor size limit. Select fewer recipients or use a smaller template.',
		);
		return false;
	}

	private async operation(
		title: LocalizationKey,
		error: MailingsLabel,
		run: (current: () => boolean) => Promise<void>,
	): Promise<void> {
		const { core } = this;
		const epoch = this.epoch;
		const generation = core.loadGeneration;
		const current = () =>
			epoch === this.epoch &&
			generation === core.loadGeneration &&
			core.element.isConnected &&
			core.host.canEditOutsideBody();
		this.busy = true;
		syncMailingsRibbon(core);
		try {
			await run(current);
		} catch (cause) {
			if (current()) {
				this.message(title, error);
				core.host.reportError(cause);
			}
		} finally {
			if (epoch === this.epoch) {
				this.busy = false;
				syncMailingsRibbon(core);
			}
		}
	}

	private message(title: LocalizationKey, text: MailingsLabel): void {
		this.show(title, (dialog) => {
			dialogMessage(dialog, text);
			dialog.actions();
		});
	}

	private show(title: LocalizationKey, build: (dialog: MailingsDialog) => void): void {
		const { element, locale } = this.core;
		if (!element.shadowRoot) return;
		this.dialog?.close();
		const dialog = createMailingsDialog(element.shadowRoot, title, locale);
		build(dialog);
		this.dialog = dialog;
		dialog.show();
	}
}
