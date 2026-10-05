# Owned binary Visio fixture

`owned-v11.vsd` is the original Apache-2.0 synthetic fixture from
ChristopherVR/ole2, commit `10668a2ba9bdc076b575290c9a84a40db305b718`.
Generator: `test/fixtures/vsd/generated.ts`; provenance:
`test/fixtures/provenance.json`. SHA-256:
`a3781bcdaadfb48f3e9669808930013197b541af164dd9e44932da6d744ef36e`.

It exercises compressed pointers, a finite page, one literal transform,
UTF-16 text, move/line geometry and opaque preserved content. Independent
libvisio callback and Windows IStorage checks establish only those scoped
assertions. It was not produced by native Visio. Microsoft Visio 16 rejects
this synthetic file as a newer or unrecognized version (HRESULT -2032466854).
It establishes parser mechanics and independent libvisio callbacks, not a
native-valid VSD file or native writer acceptance. The fixture is test-only
and is excluded from published viewer packages.
