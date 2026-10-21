/**
 * Sets a component's `displayName` and returns the component.
 *
 * A bare module-level `Component.displayName = '...'` statement is a side effect, so a host
 * bundler keeps the component (and everything it renders) even when nothing uses it. Calling this
 * helper behind a pure annotation instead lets the bundler drop an unused component together with
 * its name.
 */
export function withDisplayName<T extends { displayName?: string | undefined }>(
	component: T,
	displayName: string,
): T {
	component.displayName = displayName;
	return component;
}
