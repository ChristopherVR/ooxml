# Native line jump evidence

These drawings and measurements were generated in an owned Microsoft Visio 16
InvisibleApp COM instance on 2026-10-12. No user documents were opened.

- `native-cached-line-jumps.vsdx`: two crossing straight Dynamic connector
  instances. The horizontal connector contains a saved ArcTo jump in its local
  Geometry rows. Regenerate with `scripts/record-visio-line-jump-policy.ps1`
  using `-Output <policy.json> -FixtureOutput <drawing.vsdx>`.
- `native-line-jump-angles.json`: 160 native SVG observations of straight and
  curved connector policies at varied angles. The recorder is
  `scripts/record-visio-line-jump-angles.ps1`. Exact 45-degree ties are retained
  as observations but are not asserted as a stable rule because native
  classification changes with endpoint rounding.
- `native-line-jump-editing.json`: native SVG after this editor saves disabled
  jumps, gaps, larger arcs and reversed directions. Reproduce with
  `scripts/visio-line-jump-acceptance.ps1 -OutputDirectory <directory>`.

The accepted source-backed cache refresh is limited to straight routes with
plain supported Geometry rows. Formula dependencies and unknown row metadata
retain their saved caches and produce a diagnostic. Curved routes omit jumps
in the recorded native cases; they are not a missing line-jump feature.
