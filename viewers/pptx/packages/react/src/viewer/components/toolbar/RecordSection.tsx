import { buildRecordRibbon } from 'ooxml-ui/pptx';
import type { RibbonControlId } from 'ooxml-ui/pptx';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { WebRibbonSection } from './WebRibbonSection';

interface RecordSectionProps {
	onRecordFromBeginning: () => void;
	onRecordFromCurrent: () => void;
}
export function RecordSection(props: RecordSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const request = (id: RibbonControlId) => {
		if (id === 'record.record.fromBeginning') {
			props.onRecordFromBeginning();
		} else if (id === 'record.record.fromCurrent') {
			props.onRecordFromCurrent();
		}
	};
	return <WebRibbonSection groups={buildRecordRibbon(t)} onCommand={request} />;
}
