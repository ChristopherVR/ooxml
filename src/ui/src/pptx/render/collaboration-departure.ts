// Compatibility exports: document operations live in ooxml-core.
export {
	DEPARTURE_CHANNEL,
	removeAwarenessStatesLocally,
	createDepartureChannel,
} from 'ooxml-core/pptx/editor/render/collaboration-departure';
export type {
	DepartureNotice,
	AwarenessStatesLike,
	DepartureChannelLike,
	DepartureChannelFactory,
	DepartureChannel,
} from 'ooxml-core/pptx/editor/render/collaboration-departure';
