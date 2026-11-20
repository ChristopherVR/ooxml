// The conformance checks themselves: each one passes a conforming element and names the fault of
// an element that breaks its rule, so the gate in elements.test.ts cannot pass by checking nothing.
import { unsafeCSS } from 'lit';
import { afterAll, describe, expect, it } from 'vitest';
import { OfficeElement, controlStyles } from '../base';
import { defineElement } from '../registry';
import { CHECKS, type CheckId } from './contract';
import type { ElementClass } from './elements';
import { trackResources } from './resources';

const host = document.body.appendChild(document.createElement('div'));
const resources = trackResources(window);
afterAll(() => {
	resources.restore();
	host.remove();
});

let defined = 0;
/** Defines `ctor` under a fresh tag, through the registry unless `plain`. */
function define(
	ctor: CustomElementConstructor,
	plain = false,
): { tag: string; ctor: ElementClass } {
	const tag = `conformance-fixture-${++defined}`;
	if (plain) customElements.define(tag, ctor);
	else defineElement(customElements, tag, ctor);
	return { tag, ctor: ctor as ElementClass };
}

const run = (id: CheckId, element: { tag: string; ctor: ElementClass }, isPublic = true) =>
	CHECKS.find((check) => check.id === id)!.failure(
		{ ...element, public: isPublic },
		{
			host,
			resources,
		},
	);

/** A conforming element: declared properties, controlStyles, nothing held after removal. */
class Good extends OfficeElement {
	static override properties = { label: { type: String }, open: { type: Boolean } };
	static override styles = controlStyles(':host { color: var(--office-foreground); }');
	declare label: string;
	declare open: boolean;
	constructor() {
		super();
		this.label = '';
		this.open = false;
	}
}

describe('conformance checks', () => {
	it('pass an element that follows the contract', async () => {
		const good = define(Good);
		for (const check of CHECKS)
			expect(
				await check.failure({ ...good, public: true }, { host, resources }),
				check.id,
			).toBeUndefined();
	});

	it('name an element that is not an OfficeElement or skips the registry', async () => {
		const plain = define(class extends HTMLElement {}, true);
		expect(await run('office-element', plain)).toMatch(/another base class/);
		expect(await run('contract-stamp', plain)).toMatch(/not defined through the registry/);
		expect(await run('public-registration', define(class extends Good {}), false)).toMatch(
			/defined on use/,
		);
	});

	it('name a class field that shadows a reactive property', async () => {
		class Shadowing extends Good {
			// The fault under test: a field where a `declare` belongs.
			override label = 'field';
		}
		expect(await run('declared-properties', define(Shadowing))).toBe('own fields label');
	});

	it('name an attribute read without a property', async () => {
		class AttributeOnly extends Good {
			static override watched = ['size', 'label'];
		}
		expect(await run('attribute-properties', define(AttributeOnly))).toBe('attribute only: size');
	});

	it('name a property set before the upgrade that the element drops', async () => {
		class Forgetful extends Good {
			override connectedCallback(): void {
				super.connectedCallback();
				this.label = '';
				this.open = false;
			}
		}
		expect(await run('early-properties', define(Forgetful))).toMatch(/set before the upgrade/);
	});

	it('name styles that skip controlStyles or leave a token without its default', async () => {
		class Unstyled extends Good {
			static override styles = [unsafeCSS(':host { display: block; }')];
		}
		class BareToken extends Good {
			static override styles = [
				...controlStyles(''),
				unsafeCSS('a { color: var(--office-accent); }'),
			];
		}
		expect(await run('token-styles', define(Unstyled))).toBe('styles not built with controlStyles');
		expect(await run('token-styles', define(BareToken))).toBe(
			'tokens without a default: --office-accent',
		);
	});

	it('name what a removed element still holds', async () => {
		const release: (() => void)[] = [];
		class Leaky extends Good {
			override connectedCallback(): void {
				super.connectedCallback();
				const onEvent = (): void => {};
				const observer = new ResizeObserver(onEvent);
				window.addEventListener('resize', onEvent);
				document.addEventListener('keydown', onEvent, true);
				observer.observe(this);
				const interval = setInterval(onEvent, 60_000);
				release.push(() => {
					window.removeEventListener('resize', onEvent);
					document.removeEventListener('keydown', onEvent, true);
					observer.disconnect();
					clearInterval(interval);
				});
			}
		}
		const reason = await run('disconnect-cleanup', define(Leaky));
		for (const undo of release) undo();
		expect(reason).toMatch(/listener window:resize:false/);
		expect(reason).toMatch(/listener document:keydown:true/);
		expect(reason).toMatch(/observer ResizeObserver/);
		expect(reason).toMatch(/interval/);
	});
});
