// `ooxml-ui/bindings`: the framework-neutral kit every product's framework adapters are built on
// (Word, Excel, Visio, Teams). No framework imports and no DOM work at import time. The React, Vue
// and Solid component factories are the `/react`, `/vue` and `/solid` subpaths.
export {
	angularOutputName,
	angularOutputs,
	contractOptions,
	eventArgs,
	eventNames,
	eventOptions,
	handlerOptions,
	optionNames,
	pickProps,
	propKeys,
	readOnlyChange,
	svelteHandlerName,
	type AnyContract,
	type BindingContract,
	type EventModel,
	type EventRule,
	type ForwardPolicy,
	type MountedBinding,
	type PropRule,
} from './contract';
export {
	bindElement,
	listenEvents,
	syncClassName,
	type BindElementOptions,
	type BoundElement,
	type Unlisten,
} from './bind-element';
export { sameList, shallowEqual, type Equality } from './equality';
export {
	deferredHandle,
	handleKeys,
	installHandle,
	type DeferredHandleOptions,
	type HandleMember,
	type HandleSpec,
} from './handle';
export { createSrcLoader, nameFromUrl, type SrcLoader } from './src-loader';
export { createPropTracker, type PropTracker } from './tracker';
