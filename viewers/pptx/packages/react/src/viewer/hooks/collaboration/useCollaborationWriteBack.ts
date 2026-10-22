import { createWriteBackScheduler, writeBackReport } from 'ooxml-ui/pptx';
import { useCallback, useEffect, useMemo, useRef } from 'react';

import { buildSaveSlides } from '../../utils/template-editing';
import type { UseYjsDocumentSyncInput } from './useYjsDocumentSync';

type Input = Pick<
	UseYjsDocumentSyncInput,
	| 'doc'
	| 'config'
	| 'getSourceBytes'
	| 'getSaveOptions'
	| 'templateElementsBySlideId'
	| 'onWriteBackStatus'
>;

/** Keep framework lifecycle local and durable serialization in shared. */
export function useCollaborationWriteBack(input: Input): () => void {
	const latest = useRef(input);
	latest.current = input;
	const scheduler = useMemo(() => {
		const created = createWriteBackScheduler({
			getYDoc: () => latest.current.doc,
			getSourceBytes: () => latest.current.getSourceBytes?.() ?? null,
			getSaveOptions: () => latest.current.getSaveOptions?.(),
			getTemplateElements: () => latest.current.templateElementsBySlideId,
			mergeTemplateElements: buildSaveSlides,
			onWriteBackStatus: () =>
				latest.current.onWriteBackStatus?.(
					writeBackReport(latest.current.config, created.getState()),
				),
		});
		return created;
	}, [latest]);
	const hasWriteBack = Boolean(input.config?.onWriteBack);
	/* oxlint-disable react/exhaustive-effect-dependencies -- session identity cancels in-flight saves */
	useEffect(
		() => () => {
			scheduler.cancel();
			latest.current.onWriteBackStatus?.(undefined);
		},
		[scheduler, input.doc, input.config?.role, hasWriteBack, input.config?.externalSession],
	);
	/* oxlint-enable react/exhaustive-effect-dependencies */
	return useCallback(() => {
		const { config, doc } = latest.current;
		if (config) {
			const session = config.externalSession;
			scheduler.schedule({
				...config,
				// An inline host callback may change on an unrelated render.
				// Keep pending work, but deliver it through the current callback.
				onWriteBack: config.onWriteBack
					? (bytes) => {
							const current = latest.current;
							if (
								current.doc === doc &&
								current.config?.role === 'owner' &&
								current.config.externalSession === session
							) {
								return current.config.onWriteBack?.(bytes);
							}
						}
					: undefined,
			});
		}
	}, [scheduler, latest]);
}
