import { rebuildDrawingShapesIfCleared } from 'ooxml-ui/pptx';
import type {
	BoundingBox,
	PptxUiSmartArtTextPaneElement,
	SmartArtTextPaneChangeEvent,
} from 'ooxml-ui/pptx';
import type {
	PptxElement,
	PptxSmartArtData,
	SmartArtColorScheme,
	SmartArtStyle,
} from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { resolveSmartArtDataPalette } from '../../utils/smartart-helpers';
import { HEADING, CARD } from './inspector-pane-constants';
import { useRecentColors } from './RecentColorsContext';
import { SmartArtLayoutSwitcher } from './SmartArtLayoutSwitcher';
import { SmartArtStyleControls } from './SmartArtStyleControls';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface SmartArtPropertiesPanelProps {
	smartArtData: PptxSmartArtData;
	canEdit: boolean;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
	/** Pixel bounding box of the element, used to rebuild drawing shapes after structural edits. */
	box?: BoundingBox;
}

// --- Constants -------------------------------------------------------------

const COLOR_SCHEMES: SmartArtColorScheme[] = [
	'colorful1',
	'colorful2',
	'colorful3',
	'monochromatic1',
	'monochromatic2',
];

const STYLE_OPTIONS: SmartArtStyle[] = ['flat', 'moderate', 'intense'];

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * SmartArt inspector: layout, colours and style, and the text pane, which is the shared
 * `pptx-ui-smartart-text-pane` element (outline with levels, keyboard structure edits, move,
 * sub-item, remove and per-node style). This component feeds it the data and commits what it
 * emits through the element-update path, rebuilding the drawing when an edit cleared it.
 */
export function SmartArtPropertiesPanel({
	smartArtData,
	canEdit,
	onUpdateElement,
	box,
}: SmartArtPropertiesPanelProps): React.ReactElement {
	const { t } = useTranslation();
	const { pushColor } = useRecentColors();
	const paneRef = React.useRef<PptxUiSmartArtTextPaneElement>(null);

	const applySmartArtData = (newData: PptxSmartArtData) => {
		const reflowed = box
			? rebuildDrawingShapesIfCleared(
					newData,
					newData.layout,
					resolveSmartArtDataPalette(newData),
					newData.style ?? 'flat',
					'inspector',
					box,
				)
			: newData;
		onUpdateElement({ smartArtData: reflowed } as Partial<PptxElement>);
	};
	// The listener is attached once; it reads the latest commit path through this ref.
	const applyRef = React.useRef(applySmartArtData);
	applyRef.current = applySmartArtData;

	const updateSmartArt = (patch: Partial<PptxSmartArtData>) => {
		applySmartArtData({ ...smartArtData, ...patch });
	};

	React.useLayoutEffect(() => {
		if (paneRef.current) {
			paneRef.current.state = {
				smartArtData,
				canEdit,
				translate: (key, params) => t(key, params),
			};
		}
	}, [smartArtData, canEdit, t]);

	React.useEffect(() => {
		const pane = paneRef.current;
		if (!pane) {
			return;
		}
		const commit = (event: Event) => {
			const { smartArtData: next, recentColor } = (event as SmartArtTextPaneChangeEvent).detail;
			applyRef.current(next);
			if (recentColor) {
				pushColor(recentColor);
			}
		};
		pane.addEventListener('pptx-smartart-data-change', commit);
		return () => pane.removeEventListener('pptx-smartart-data-change', commit);
	}, [pushColor]);

	return (
		<div className={CARD} role='group' aria-label={t('pptx.smartart.title')}>
			<div className={HEADING}>{t('pptx.smartart.title')}</div>
			<div className='space-y-2'>
				<SmartArtLayoutSwitcher
					smartArtData={smartArtData}
					canEdit={canEdit}
					onUpdateSmartArt={updateSmartArt}
				/>

				<SmartArtStyleControls
					colorSchemes={COLOR_SCHEMES}
					styleOptions={STYLE_OPTIONS}
					colorScheme={smartArtData.colorScheme ?? 'colorful1'}
					style={smartArtData.style ?? 'flat'}
					canEdit={canEdit}
					onChangeColorScheme={(colorScheme) => updateSmartArt({ colorScheme })}
					onChangeStyle={(style) => updateSmartArt({ style })}
				/>

				<pptx-ui-smartart-text-pane ref={paneRef} data-testid='smartart-text-pane' />
			</div>
		</div>
	);
}
