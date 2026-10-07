// Compatibility exports: document operations live in ooxml-core.
export {
	getUserInitials,
	buildActiveSessionUsers,
} from 'ooxml-core/pptx/editor/render/collaboration-active-session';
export type {
	ActiveSessionRemoteUserInput,
	ActiveSessionUserDescriptor,
	BuildActiveSessionUsersParams,
} from 'ooxml-core/pptx/editor/render/collaboration-active-session';
