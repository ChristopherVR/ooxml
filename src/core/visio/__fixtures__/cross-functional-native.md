# Native cross-functional flowchart

`cross-functional-native.vsdx` is the diagram Visio 16 builds from its own Cross-Functional
Flowchart template (`NEW_XFUNC_M.VSTX`, driven through COM in an invisible instance and saved
with `SaveAs`): the CFF Container (sheet 1), the Swimlane List (4) with two Swimlanes (6, 9),
the Phase List (5) with one Separator (12). Read through COM, `ContainerProperties.GetMemberShapes`
of the container is 4 and 5, `GetListMembers` of the Swimlane List is 9 and 6 (top lane first),
and the Relationships formulas use the kinds the editor now writes: 1 container members, 4
member of, 2 list members, 5 list member of, 11 the sibling list.

The `cross-functional-native-*.vsdx` files are what the same Visio saved after one change each to
that chart, made through COM with `EventsEnabled` set so the Cross-Functional add-in reacts:
`-phase` after dropping a Separator master at x = 8 in (the add-in splits the phase and lists
the new separator first), `-notitle`, `-nophases` and `-style3` after setting `visShowTitle`,
`visShowPhase` and `visCFFStyle` on the frame, `-rotated` after setting `visRotateLabel` on the
Swimlane List. `edit-swimlane-cff.test.ts` replays the same changes with the editor and compares
every cell Visio changed. The ribbon's Direction command (add-in `/cmd=4`, an orientation dialog)
blocks an automated Visio and was not run.

The editor's own output (no masters, the same User rows and Relationships formulas as direct
cells) was opened the same way; Visio reported the same container, list and lane structure for it.
No stencil content was edited or copied into the runtime; the file only checks that the reader
recognises what Visio itself wrote.
