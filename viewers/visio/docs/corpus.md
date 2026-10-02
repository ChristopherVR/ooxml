# Reproduce external corpus checks

The optional corpus runner checks 19 real drawings and 10 malformed security inputs from two primary upstream repositories. It never downloads, uploads or redistributes fixture bytes. Each input must match the size and SHA-256 recorded in `tests/corpus-baseline.json`.

- [LibreOffice/libvisio](https://github.com/LibreOffice/libvisio/tree/49fb9d3a9d21d4374cad782925e48c577a41f5be/src/test/data), revision `49fb9d3a9d21d4374cad782925e48c577a41f5be`: 14 real VSDX files and two malformed inputs. The repository declares MPL-2.0 and includes the fixtures in its test distribution.
- [Apache POI](https://github.com/apache/poi/tree/732120980140d5ed64b482c470e0b625cdb1ab15/test-data/diagram), revision `732120980140d5ed64b482c470e0b625cdb1ab15`: five real VSDX files and eight malformed inputs. The repository provides Apache-2.0 license and notice files.

Retain upstream provenance, licenses and notices. Confirm fixture-specific content rights before product redistribution. This viewer repository includes factual source identifiers, hashes and regression expectations only.

## Run

Prepare local upstream source checkouts containing the paths listed in the manifest. After building the core and viewer:

```sh
npm run test:corpus -- /path/to/libvisio /path/to/poi
```

Append `--svg` to additionally serialize each page in a simulated DOM and verify safe XML, local-only resource references, retained diagnostics and absence of viewer controls or event attributes. All 22 current pages pass this structural export check; it is not a browser or pixel-equivalence test.

The runner rejects missing or changed bytes. It checks structured rejection codes for malformed packages, normalized page/shape/path counts, defensive scene validation, selected upstream fill-color assertions and five source-confirmed rounded rectangles. It reports 22 pages, 706 normalized shapes and 627 paths for the current pinned baseline. Changes to expected counts require source review rather than automatic acceptance.

## Meaning and limits

Normalized counts guard against implementation regressions; they are not native Visio reference values. Raw XML counts differ because deleted shapes are removed and inherited master children can be added. The malformed inputs are excluded from fidelity denominators.

The selected fill expectations come from [libvisio's import tests](https://github.com/LibreOffice/libvisio/blob/49fb9d3a9d21d4374cad782925e48c577a41f5be/src/test/importtest.cpp). They do not certify every theme effect. Saved rectangle radii are verified against package XML.

Embedded EMF thumbnails are partial, sometimes stale references. Secondary rasterization can reveal gross omissions but cannot establish fixed-font, full-page or native Microsoft Visio parity. See [the verification record](verification.md) for exclusions and remaining evidence gaps.
