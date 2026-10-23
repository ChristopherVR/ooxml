# Framework binding kit (`ooxml-ui/bindings`)

The code every product's framework adapters share: a product declares its element as a
contract, and the kit wires the element (events, forwarded props, the imperative handle) once for
every framework. Nothing here imports a framework or touches the DOM at import time. The summary
for users of the package is in `src/ui/README.md` ("Framework binding kit"); this page is for
the product that defines a contract and the adapter that consumes it.

| Module                           | Exports                                                                                                                                                                                                                                |
| -------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `contract.ts`                    | `BindingContract`, `PropRule`, `EventRule`, `EventModel`, `ForwardPolicy`, `eventArgs`, `readOnlyChange`, `pickProps`, `eventOptions`, `handlerOptions`, `contractOptions`, `angularOutputs`, `svelteHandlerName`, `angularOutputName` |
| `bind-element.ts`                | `bindElement`, `listenEvents`, `syncClassName`                                                                                                                                                                                         |
| `tracker.ts`                     | `createPropTracker`                                                                                                                                                                                                                    |
| `equality.ts`                    | `shallowEqual`, `sameList`                                                                                                                                                                                                             |
| `handle.ts`                      | `deferredHandle`, `installHandle`, `handleKeys`                                                                                                                                                                                        |
| `src-loader.ts`                  | `createSrcLoader`, `nameFromUrl`                                                                                                                                                                                                       |
| `react.ts`, `vue.ts`, `solid.ts` | `createReactBinding`, `defineVueBinding`, `createSolidBinding` (the `ooxml-ui/bindings/react`, `/vue`, `/solid` subpaths)                                                                                                              |

React, Vue and Solid are optional peer dependencies, so a product that never imports a factory
does not need the framework installed. Angular components and Svelte components need their own
compilers, so each product keeps a small adapter for them, written over the same helpers.

## 1. Define the contract

A product's contract is one object of type `BindingContract<E, P, H, N, O>`:

- `E` is the element, `P` the props a framework passes, `H` the handle, `N` the element event
  names and `O` the handler option names.
- `props` lists every forwarded prop, in forwarding order. Each rule has a `forward` policy:
  `always` (written on every update), `changed` (written when the value changed since it was last
  written), `defined` (like `changed`, but an `undefined` prop is skipped, so a user's choice inside
  the element survives a re-render, for example a renamed file), or `manual` (the product's mount
  function writes it, for ordering or supersession). A rule may also give a `fallback`, a `same`
  comparison, an `initial` value counted as already written, and an `apply` function for a
  property that is not a plain assignment.
- `events` maps each element event to its handler option and arguments. `eventArgs.detail` is the
  default; `eventArgs.event` passes the event itself for a cancelable event; `eventArgs.field(name)`
  passes one field of the detail. A two-way prop is declared with a `model` (`readOnlyChange`
  mirrors the host's own switch, never a collaboration lock).
- `handle` has one entry per member of the handle: `get` (a property that throws before mount),
  `call` (a method that throws), `async` (a method that rejects), `optional-call` (a method that does
  nothing before mount) or `{ fallback }` (a property that reads the fallback). A member missing from
  `handle` is a compile error.

Word's contract is the reference (`viewers/docx/packages/bindings/src/contract.ts`,
`EDITOR_CONTRACT`); its `contract-parity.test.ts` checks every adapter against it. Excel, Visio
and Teams follow the same shape (`viewers/xlsx/packages/bindings/src/contract.ts`,
`viewers/visio/packages/bindings/src/common.ts` (`VIEWER_CONTRACT`), `viewers/teams/packages/web-component/src/bindings-contract.test.ts`).

## 2. Wire the element

The product's mount function creates the element, then calls the kit:

```ts
const bound = bindElement(element, EDITOR_CONTRACT, () => options, {
	onEvent: (name) => recordWhatTheElementEmitted(name),
});
bound.update(options); // forwards every prop by its policy; `manual` props are skipped
// on teardown: bound.unlisten(); element.remove();
```

`bindElement` listens to every event in the contract and calls the handler option read at event
time (so the latest props win without re-binding). `update` forwards the props that changed.
Anything the kit does not own stays in the product's mount function around `update`: Word's
document model after collaboration, Excel's workbook bytes, Visio's revision-aware patches, the
`src` supersession. For a `src` prop, `createSrcLoader(load, report)` fetches the URL, names the
file from its last segment and drops a fetch that a newer `src`, new bytes or unmounting has
superseded.

## 3. Adapt it to each framework

- **React**: `createReactBinding({ displayName, handle, mount, options, host })` returns a
  `forwardRef` component. It mounts once, forwards each render's props through `update` and
  exposes the deferred handle as its `ref`.
- **Vue**: `defineVueBinding({ name, props, contract, handle, mount, callbacks, options })` declares
  the props, kebab-case emits for every element event, `v-model` for modelled events and the handle
  on the template ref (`expose`). Vue's template refs read `get` members leniently before mount
  by default (`lenientGetters`).
- **Solid**: `createSolidBinding({ handle, mount, options })` returns a function component that
  creates its host `<div>`, re-applies options reactively and hands the handle to the `editorRef`
  callback prop (`ref` names another).
- **Angular**: `installHandle(prototype, spec, current)` defines the handle members on the component
  class, and `angularOutputs(contract, component)` emits each event as the output
  `angularOutputName(option)` (`onDocumentChange` becomes `documentChange`), with the modelled
  output for `readOnly`.
- **Svelte**: `svelteHandlerName(option)` gives the lower-case callback prop name
  (`onDocumentChange` becomes `ondocumentchange`); the component's exports are the handle.
- **Vanilla**: call `bindElement` directly and return `{ update, destroy }` (the `MountedBinding`
  type).

`deferredHandle(spec, current, options)` builds the handle every adapter returns: each member
reads the binding that is mounted at call time, so one handle object serves the whole life of the
component. Before mount, `get` members throw the product's `message` option (Word uses
`Editor is not mounted`, unless `lenientGetters` is set), `call` members throw, `async` members reject, `optional-call` members do nothing and fallbacks
read their value.

## 4. Tests

Each product has a parity test that reads its contract and checks every adapter's props, events and
handle members against it, so an adapter cannot drift from the contract. A change to the kit
itself is covered by `bindings.test.ts` (policies, handlers, contract mapping, the handle and the
source loader) and `frameworks.test.ts` (the React, Vue and Solid factories) in this folder.
