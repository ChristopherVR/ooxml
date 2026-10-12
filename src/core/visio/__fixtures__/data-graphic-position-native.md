# Generated graphic position acceptance

`scripts/visio-data-graphic-position-acceptance.ps1` creates a synthetic rectangle
with text, bar and icon graphics using the core. Generated local 2D graphics use
`Sheet.ID!PinX/PinY` formulas retaining their initial offsets.

An owned Microsoft Visio 16 COM instance opened the authored and core-moved files
on 2026-10-12. Moving the authored owner in Visio and reopening the core-moved
version both placed all four graphic parts at the expected coordinates. Eight
observations are recorded in `data-graphic-position-native.json`.

These remain separate generated shapes, not native data-graphic masters. Group
and stencil owners retain fixed positions. Pin offsets follow translations;
resizing, rotation-relative positioning, and automatic value refresh are outside
this acceptance. The Delete workflow includes generated followers in the same undoable operation; direct core deletion refuses an owner when its followers are not included.
