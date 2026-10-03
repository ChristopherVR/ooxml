# Legacy VSD preview

`loadVisio(bytes, options)` detects VSD compound files and VSDX ZIP packages from
their container bytes. VSDX keeps the existing `parseVsdx` behavior and options.
Binary VSD decoding belongs to `@christophervr/ole2`; this adapter only creates a
Visio scene from its structured model.

```ts
import { loadVisio } from 'ooxml-core/visio';

const scene = await loadVisio(bytes);
console.log(scene.format); // 'vsd' or 'vsdx'
console.log(scene.diagnostics);
```

The VSD preview currently accepts version 11 drawings with explicit page size,
unit drawing scale, independent shape transforms, and move/line geometry. It
uses the existing Visio inch coordinates and transform implementation. Groups,
foreign objects, parent/master dependencies, unsupported geometry, unresolved
scale, missing dimensions and text field controls fail explicitly.

Legacy styles, visibility, layers, connections, background assignment and text
block formatting remain unresolved. Accepted scenes include diagnostics for a
transparent fill, black outline and plain 12pt Arial text fallback. Stored text
is inert; the importer never executes macros or embedded objects. It does not
assert native Visio visual fidelity.

The legacy scene is a preview, not VSDX export. Keep `scene.format` when presenting
save controls. `editVsdx` requires an original VSDX package and does not convert a
VSD compound file. The shared `ole2` model separately owns supported binary
edits and preservation of opaque streams.

`VisioDocument.format` now has type `'vsdx' | 'vsd'`. Existing callers assigning
it to the literal type `'vsdx'` must first check `scene.format === 'vsdx'`.
The `parseVsdx` function still accepts only VSDX packages and retains its behavior.

Input, shape, geometry, text and diagnostic budgets apply. Legacy `limits` accepts
`maxInputBytes` and `maxRuntimeMs`; ZIP-entry and XML limits apply to VSDX only.
The shared decoder also bounds binary traversal and decompression. The elapsed-time check surrounds
synchronous decoding; callers requiring hard preemption should use a worker.
