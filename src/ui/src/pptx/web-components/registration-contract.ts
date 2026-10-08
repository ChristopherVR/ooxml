/**
 * pptx layer over the shared registry contract in `../../registry`. The shared
 * `office-ui.web-control-contract` stamp is the implementation; the legacy
 * `pptx-viewer.web-control-contract` symbol is still written and read so a
 * window that mixes an older pptx bundle keeps failing loudly.
 */
import { assertContract, CONTRACT_REVISION, markContract } from '../../registry';

const LEGACY_KEY = Symbol.for('pptx-viewer.web-control-contract');

type LegacyConstructor = CustomElementConstructor & { [LEGACY_KEY]?: number };

/** Preflight every tag before installing styles or defining any new controls. */
export function assertWebControlContract(
	registry: Pick<CustomElementRegistry, 'get'>,
	names: readonly string[],
): void {
	for (const name of names) {
		assertContract(registry, name);
		const revision = (registry.get(name) as LegacyConstructor | undefined)?.[LEGACY_KEY];
		if (revision !== undefined && revision !== CONTRACT_REVISION) {
			throw new Error(
				`Incompatible ${name} contract (${revision}; expected ${CONTRACT_REVISION}). Use matching viewer bindings in this window.`,
			);
		}
	}
}

/** Only stamp implementations defined by this bundle, never unmarked legacy tags. */
export function markWebControlContract(ctor: CustomElementConstructor): void {
	markContract(ctor);
	Object.defineProperty(ctor, LEGACY_KEY, { value: CONTRACT_REVISION });
}
