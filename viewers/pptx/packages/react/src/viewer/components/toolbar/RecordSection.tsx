import { buildRecordRibbon, runRecordRibbonAction } from 'ooxml-ui/pptx';
import type { RecordTarget, RibbonControlId } from 'ooxml-ui/pptx';
import React, { useRef } from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { WebRibbonSection } from './WebRibbonSection';

interface RecordSectionProps {
	onRecordFromBeginning: () => void;
	onRecordFromCurrent: () => void;
	getTarget?: (() => RecordTarget | undefined) | undefined;
}
export function RecordSection(props: RecordSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const live = useRef(props.getTarget);
	live.current = props.getTarget;
	const request = (id: RibbonControlId, host: HTMLElement) => {
		if (id === 'record.record.fromBeginning') {
			props.onRecordFromBeginning();
		} else if (id === 'record.record.fromCurrent') {
			props.onRecordFromCurrent();
		} else runRecordRibbonAction(host, id, () => live.current?.(), t);
	};
	return (
		<WebRibbonSection groups={buildRecordRibbon(t, props.getTarget?.())} onCommand={request} />
	);
}
