# Native print count evidence

Owned synthetic Microsoft Visio 16.0.20430.20146 COM probes, measured against the read-only Microsoft Print to PDF context. The original 209-entry evidence archive and separate 244-entry extension remain immutable; their SHA-256 values are recorded in print-count-native.json. The 79 raw VSDX sources retain their original bytes and individual hashes.

The bounded core profile admits 66 cases, refuses five recorded near-boundary cases as boundary-precision, and excludes eight fit, centering or unequal-scale cases. The five precision refusals retain their observed native counts; no rounding epsilon repairs them. Physical extents include the existing drawing-scale projection. Counts do not establish tile origins, clips, printer output, or cross-driver behavior.

Native PDF exports each produced one full physical drawing page. They do not establish printer tiling. Native reopen automation stalled, so native reopened acceptance is zero. Core tests check package no-op identity and unrelated page rename preservation, not native reopen acceptance.

Defaults resolve only from the unique canonical No Style0 static cells with measured page/document style references0. Direct DocumentSheet print cells are ineffective. Explicit page overrides win; custom/chained styles, malformed/dynamic local cells and duplicates are refused. Count-only metadata never changes legacy page-break geometry or editing permissions.
