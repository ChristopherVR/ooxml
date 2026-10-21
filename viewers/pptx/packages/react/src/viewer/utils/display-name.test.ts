import { describe, expect, it } from 'vitest';

import { HorizontalRuler, VerticalRuler } from '../components/canvas/RulerStrips';
import { ConnectorElementRenderer } from '../components/elements/ConnectorElementRenderer';
import { Extrusion3DOverlay } from '../components/elements/Extrusion3DOverlay';
import { StaticElementRenderer } from '../components/StaticElementRenderer';
import { PowerPointViewer } from '../PowerPointViewer';
import { withDisplayName } from './display-name';

describe('withDisplayName', () => {
	it('sets the display name on the component it returns', () => {
		const component: { displayName?: string } = {};
		expect(withDisplayName(component, 'Named')).toBe(component);
		expect(component.displayName).toBe('Named');
	});

	it('keeps the display names of the components defined through it', () => {
		expect(PowerPointViewer.displayName).toBe('PowerPointViewer');
		expect(StaticElementRenderer.displayName).toBe('StaticElementRenderer');
		expect(ConnectorElementRenderer.displayName).toBe('ConnectorElementRenderer');
		expect(Extrusion3DOverlay.displayName).toBe('Extrusion3DOverlay');
		expect(HorizontalRuler.displayName).toBe('HorizontalRuler');
		expect(VerticalRuler.displayName).toBe('VerticalRuler');
	});
});
