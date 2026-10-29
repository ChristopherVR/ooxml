# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs releases whenever any published file changes, not only on
conventional commits.

## [1.33.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.33.0) - 2026-10-29

### Features

- **docx:** Stamp the last-modified time when a host asks on save ([8fd056a](https://github.com/ChristopherVR/ooxml/commit/8fd056a79271bb55195cd0f500da876e59a83393))
- **docx:** Save edits to the formatting of existing paragraph styles ([4716ef0](https://github.com/ChristopherVR/ooxml/commit/4716ef0dcbcdb50b47b3823b703a7231630538e4))
- **teams:** Lock the app when the workspace schema is newer ([0ef86c8](https://github.com/ChristopherVR/ooxml/commit/0ef86c8ad19a77ef3d277a9273fc531ebac0c0d2))
- **xlsx:** Add forecast sheet ([cacb63e](https://github.com/ChristopherVR/ooxml/commit/cacb63e5e7ba8e3e42ecd2af5045a636eedfd41b))
- **teams:** Offer moderation in the app and back off on refusals ([0801ef2](https://github.com/ChristopherVR/ooxml/commit/0801ef2f68dc98e0b30a04687708b6bfdd5bf549))
- **docx:** Stamp the host review author as the last modifier on save ([fa39dbe](https://github.com/ChristopherVR/ooxml/commit/fa39dbe0e4d60fd6cb56f64a02dedba59678292a))
- **ui:** Merge, split and edit rows and columns of merged tables ([2dadb31](https://github.com/ChristopherVR/ooxml/commit/2dadb312c052ade9a0be0e55b688df3e821daab1))
- **xlsx:** Draw worksheet shapes and form controls ([656762d](https://github.com/ChristopherVR/ooxml/commit/656762d445d8516c5415ceaa0a6036e194d603f5))
- **xlsx:** Show header pictures, offsets and theme colours in print ([47fb4bd](https://github.com/ChristopherVR/ooxml/commit/47fb4bdef22f45959d68473248b31ed779a92d6b))
- **xlsx:** Add scenario summary and merge ([58d9b07](https://github.com/ChristopherVR/ooxml/commit/58d9b074d052a1f2a82e36ca700b785057ba48a7))
- **xlsx:** Add custom views ([671a7a9](https://github.com/ChristopherVR/ooxml/commit/671a7a965bc1bf5d877b0f89a2f3679c2ca86a69))
- **xlsx:** Group and ungroup sparklines ([4bbd9f8](https://github.com/ChristopherVR/ooxml/commit/4bbd9f8ca19b9570fc8ff11b0347164526f46926))
- **xlsx:** Copy multi-area selections and lift the paste tile cap ([d60224f](https://github.com/ChristopherVR/ooxml/commit/d60224f1684f8af78848ca6372b29c05f6c6830d))
- **xlsx:** Paste into multi-area selections and refuse multi-area cut ([15061b1](https://github.com/ChristopherVR/ooxml/commit/15061b1dec4903a861fc1ef700247d3fda9ae50b))

### Bug Fixes

- **ui:** Allow a table row covered by vertical merges ([c95178b](https://github.com/ChristopherVR/ooxml/commit/c95178b7fce1574dc16dd3540c7a8c18c26cb898))
- **xlsx:** Activate the new sheet after forecast and scenario summary ([2132fd7](https://github.com/ChristopherVR/ooxml/commit/2132fd79c14921b8555a2f4251e142654bc7b05d))
- **docx:** Size merged table rows and spaces like Word ([853787f](https://github.com/ChristopherVR/ooxml/commit/853787f8c59c6be78313ba1c8caf464f1d13501d))

### Testing

- **teams:** Cover moderator actions in the mounted app ([f2fef2f](https://github.com/ChristopherVR/ooxml/commit/f2fef2fb16a73ca7321647438ce4390e77321fd8))
- **ui:** Expect merge and split in the table context menu ([91bc2de](https://github.com/ChristopherVR/ooxml/commit/91bc2dec8a926702371d89804a41d79a87d6d317))

### Styling

- Drop unused visio imports and format merged files ([41d4e2a](https://github.com/ChristopherVR/ooxml/commit/41d4e2a2bbd5b84e9c3a66e66dc6923772b93bbc))

### Chores

- **ui:** Regenerate the custom elements manifest ([0818fcb](https://github.com/ChristopherVR/ooxml/commit/0818fcba41fd1a4b9535b64c81f015cf56ef07ff))

## [1.32.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.32.0) - 2026-10-28

### Features

- **pptx:** Add opt-in whole-list conversion with listConversionScope ([#229](https://github.com/ChristopherVR/ooxml/issues/229)) ([bf7be3e](https://github.com/ChristopherVR/ooxml/commit/bf7be3e5cedd7ff9ed9e6cc62b0bd5be7ec363c4))

### Bug Fixes

- **pptx:** Stretch textPlain WordArt to fill its shape ([#226](https://github.com/ChristopherVR/ooxml/issues/226)) ([120d3ce](https://github.com/ChristopherVR/ooxml/commit/120d3ce4e4e4dd8c5830902230fabb5e2c8dcfe7))

## [1.31.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.31.0) - 2026-10-28

### Features

- **pptx:** Expose the slide thumbnail rail as a building block in every binding ([#228](https://github.com/ChristopherVR/ooxml/issues/228)) ([2b77f02](https://github.com/ChristopherVR/ooxml/commit/2b77f026e1e26e6a5fe3c52f275608f39bc0ea8e))
- **text:** Share font selection with word script layout ([59ad0a2](https://github.com/ChristopherVR/ooxml/commit/59ad0a2e8249d3ffc9be34c2c9e6f9ffe8b3b07a))
- **chart:** Share manual plot layout across chart painters ([18b79c0](https://github.com/ChristopherVR/ooxml/commit/18b79c04bb1d2aa914630453413d995d0bb54922))
- **chart:** Share smooth series paths and authored strokes ([7329e5c](https://github.com/ChristopherVR/ooxml/commit/7329e5cbb546a041c0e67417fbe84e9d3dc415a2))
- **chart:** Share point label offsets and pie leader lines ([861f78f](https://github.com/ChristopherVR/ooxml/commit/861f78f5d9ba0fa93ed6e85ddfc209fd2bf828ca))
- **drawingml:** Render guide-driven custom shapes in word ([81558e6](https://github.com/ChristopherVR/ooxml/commit/81558e6fcfdb2a64ced22661410cbdfba954cea2))
- **drawingml:** Share picture crop with excel ([f093c0f](https://github.com/ChristopherVR/ooxml/commit/f093c0f6d1c2ea96736613990e870143d770774e))
- **ui:** Share live colour transactions with excel fill previews ([6a3415b](https://github.com/ChristopherVR/ooxml/commit/6a3415b97f754476641fdb51b45b4ccf462c9ebc))
- **chart:** Share calendar date axis rendering ([474de5c](https://github.com/ChristopherVR/ooxml/commit/474de5c4f2418ee7b67ec34a5e08c8eefa4bb992))
- **chart:** Share chart data table layout and rendering ([09fbdf1](https://github.com/ChristopherVR/ooxml/commit/09fbdf1ec8016b376b67053d8a5575fc3d1ef145))
- **drawingml:** Share ordered picture effects with excel ([b669575](https://github.com/ChristopherVR/ooxml/commit/b669575a041ab25169a7c83757075fe9dffe390a))
- **chart:** Share funnel geometry and neutral rendering ([28f6ce3](https://github.com/ChristopherVR/ooxml/commit/28f6ce302fa5791a11819209d6501e97e971b80b))
- **chart:** Share box distributions and neutral rendering ([6390e30](https://github.com/ChristopherVR/ooxml/commit/6390e30ac6994e6891f8324e4264949f66e0a0a0))
- **diagram:** Paint cached pictures in word and excel ([f254d51](https://github.com/ChristopherVR/ooxml/commit/f254d51eeed81279dd9f971cb7ac496364738ec2))
- **ui:** Expose shared preview and recovery helpers ([af92a87](https://github.com/ChristopherVR/ooxml/commit/af92a87969118bef3f07a1a37412992dad4f652f))
- **drawingml:** Share source-backed text body 3d codecs ([bcf2a0a](https://github.com/ChristopherVR/ooxml/commit/bcf2a0a491c6e4dc4e5fd4f5b615f3673285cadf))
- **pptx:** Add an /ai entry and createAiAssistant to every binding ([#219](https://github.com/ChristopherVR/ooxml/issues/219)) ([ecd4d43](https://github.com/ChristopherVR/ooxml/commit/ecd4d43acfc0133c2f9c1e3320f366244ba3bba3))
- **visio:** Improve selection dragging and resizing ([60cf8dd](https://github.com/ChristopherVR/ooxml/commit/60cf8dd9c547f16e114b98ae01fed0968aca5866))

### Bug Fixes

- **pptx:** Scope collaboration undo to the local participant ([55b8bad](https://github.com/ChristopherVR/ooxml/commit/55b8badcf0ec06a7294debff8deff844dcfafb62))
- **pptx:** Autocorrect only the text typed in this edit ([0e5a85e](https://github.com/ChristopherVR/ooxml/commit/0e5a85e496e0d53e36c00e6fdfd5dd3749337d8a))
- **ui:** Order visio recovery cleanup after pending writes ([2fc797e](https://github.com/ChristopherVR/ooxml/commit/2fc797ecb87e72e6fcfe2f72aba50b7eede55646))
- **ui:** Preserve injected recovery timer receivers ([2ca6ec8](https://github.com/ChristopherVR/ooxml/commit/2ca6ec8db4ee0c783bc44abc647e233430a6b19c))
- **drawingml:** Bound picture luminance and report extensions ([a36e847](https://github.com/ChristopherVR/ooxml/commit/a36e84725c483815ae7508872dd4d89c5a8375b2))
- **ui:** Preserve the custom font registration result contract ([2b4876c](https://github.com/ChristopherVR/ooxml/commit/2b4876c5a02cb11e598cf84ccf921c1d870a22a9))
- **chart:** Retain table names and exclude hidden box series ([292dc25](https://github.com/ChristopherVR/ooxml/commit/292dc25c22544d9179d5c0fd31327a3d0aafb462))
- **chart:** Match native funnel label and box mean defaults ([06bbf26](https://github.com/ChristopherVR/ooxml/commit/06bbf26e496f00270eca67472087454496217cc1))
- **drawingml:** Preserve anchors and guard text 3d edits ([617a029](https://github.com/ChristopherVR/ooxml/commit/617a02979d5cdae5e3818afd8fb79af367de25eb))
- **ui:** Refresh SmartArt picture resolver manifest ([94ff626](https://github.com/ChristopherVR/ooxml/commit/94ff626417a988f4caf76c897182b10e1a47b6bc))

### Refactor

- **ui:** Share font catalogues and browser font services ([0b08f74](https://github.com/ChristopherVR/ooxml/commit/0b08f74eaa4c5bb763ec90f0708d8ea6f62368f8))
- **ui:** Share modal command sessions in word and excel ([b90ad2d](https://github.com/ChristopherVR/ooxml/commit/b90ad2d9b33d4c3df44b286e2922a70273ff0b0b))
- **ui:** Share recovery scheduling across powerpoint and visio ([96d0ab2](https://github.com/ChristopherVR/ooxml/commit/96d0ab2e77c41447ef5ca3aacc15f1e0d36ad5f2))
- **opc:** Share hyperlink policy with explicit host protocols ([5f8ac72](https://github.com/ChristopherVR/ooxml/commit/5f8ac72edbe4fff4d2fa303ed3ce8a471762bd12))

### Chores

- Integrate current main for reuse verification ([a22d5f5](https://github.com/ChristopherVR/ooxml/commit/a22d5f53406330909d3853ab64a4d850bc3c0ad1))
- Integrate main before shared capability push ([9dc78da](https://github.com/ChristopherVR/ooxml/commit/9dc78da5262a854807b9fc8395f1fd97f71c2eb4))

## [1.30.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.30.0) - 2026-10-28

### Features

- **diagram:** Share interpreted SmartArt previews across products ([f9dd019](https://github.com/ChristopherVR/ooxml/commit/f9dd0197da1871da3178a2d3ee880fe27266f0c0))
- **diagram:** Share SmartArt style and color authoring ([86b43a4](https://github.com/ChristopherVR/ooxml/commit/86b43a497b5effa071b25e73dab91ab4bb18c296))
- **ui:** Edit ChartEx charts in Word and Excel ([35ca2bd](https://github.com/ChristopherVR/ooxml/commit/35ca2bd824525a67220c533b6487b604118526b2))
- **ui:** Expose shared advanced chart format controls ([ad2513b](https://github.com/ChristopherVR/ooxml/commit/ad2513b72b76c554d789dfd96f0a4968d4553eaa))
- **diagram:** Attach generated caches in Word and Excel ([0ba0947](https://github.com/ChristopherVR/ooxml/commit/0ba09477b7c5a6b251dd607d0775feec4655abf0))
- **ui:** Expose shared advanced ChartEx formatting ([1346009](https://github.com/ChristopherVR/ooxml/commit/13460090ee3104f6e2d7cd492c163e5910d965ec))

### Bug Fixes

- **chart:** Preserve source content across chart consumers ([1ce74e9](https://github.com/ChristopherVR/ooxml/commit/1ce74e92f8864a046f3ef756b35bf8568c685c5b))
- **ui:** Localize unsupported chart diagnostics ([ce1c25b](https://github.com/ChristopherVR/ooxml/commit/ce1c25b09a209bde41cfffc7f0104e028f3f15a0))
- **diagram:** Report partial interpreted preview capabilities ([ea9e474](https://github.com/ChristopherVR/ooxml/commit/ea9e4746aa826a00a8bc49ac58852687e62eaf12))
- **pptx:** Morph offstage text size changes and keep run alpha ([75749c0](https://github.com/ChristopherVR/ooxml/commit/75749c0618fe24714d9f086ca7248f9f3e41e729))
- **pptx:** End inline edit on Escape and scope collab undo origin ([1338d38](https://github.com/ChristopherVR/ooxml/commit/1338d388f137f8542ce07dbd61f5b6186a842d41))
- **pptx:** Keep exported text in place when page stylesheets load late ([#227](https://github.com/ChristopherVR/ooxml/issues/227)) ([0b378f4](https://github.com/ChristopherVR/ooxml/commit/0b378f4e3c709c57e96a63394155c7e6467b4cfa))
- **pptx:** Converge when two clients seed an empty room ([#217](https://github.com/ChristopherVR/ooxml/issues/217)) ([ee06f53](https://github.com/ChristopherVR/ooxml/commit/ee06f5389a9f277704640e4efaa567b3d5718507))

## [1.29.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.29.1) - 2026-10-26

### Bug Fixes

- **angular:** Load the assistant runtime on demand ([ededbfe](https://github.com/ChristopherVR/ooxml/commit/ededbfe953532628ace701c2be19310212509d1d))

## [1.29.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.29.0) - 2026-10-26

### Features

- **ui:** Add a lightweight pptx AI host entry ([#210](https://github.com/ChristopherVR/ooxml/issues/210)) ([d47a350](https://github.com/ChristopherVR/ooxml/commit/d47a350083895e3c07294eb2b6e8819102ee8599))
- **pptx:** Let hosts bound deck loads with limits ([#216](https://github.com/ChristopherVR/ooxml/issues/216)) ([58a5540](https://github.com/ChristopherVR/ooxml/commit/58a5540075982935d91cbc9c97a5c5d66359a180))
- **pptx:** Decode embedded fonts on first use ([#209](https://github.com/ChristopherVR/ooxml/issues/209)) ([a1c313a](https://github.com/ChristopherVR/ooxml/commit/a1c313ab23421786f632f95927735a7c2e89bbb1))

### Performance

- **pptx:** Load the outline font parser on demand ([#208](https://github.com/ChristopherVR/ooxml/issues/208)) ([7b930b3](https://github.com/ChristopherVR/ooxml/commit/7b930b3fc57bcb2ff1047bed1c6a25bb6484127c))

### Dependencies

- **deps:** Bump production minor and patch dependencies ([#213](https://github.com/ChristopherVR/ooxml/issues/213)) ([3cfdd01](https://github.com/ChristopherVR/ooxml/commit/3cfdd016aa61cacb0540491c5ab13974e32f8e1e))

### Chores

- **deps-dev:** Bump development minor and patch dependencies ([#215](https://github.com/ChristopherVR/ooxml/issues/215)) ([a445b3e](https://github.com/ChristopherVR/ooxml/commit/a445b3e3cc2191341063c799768c5aad13981751))

## [1.28.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.28.1) - 2026-10-26

### Bug Fixes

- **pptx:** Keep comment @-mentions on their text through edits, trims and saves ([#206](https://github.com/ChristopherVR/ooxml/issues/206)) ([09a420f](https://github.com/ChristopherVR/ooxml/commit/09a420fb44364fff04df9aa5550684519bec02cd))

## [1.28.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.28.0) - 2026-10-25

### Features

- **pptx:** Let hosts block external content with allowExternalContent ([#201](https://github.com/ChristopherVR/ooxml/issues/201)) ([903c8e2](https://github.com/ChristopherVR/ooxml/commit/903c8e201dbd36252a5e82a388dcb59e30f3afc4))

## [1.27.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.27.1) - 2026-10-25

### Bug Fixes

- **pptx:** Preview colour drags as one undo step ([#203](https://github.com/ChristopherVR/ooxml/issues/203)) ([27cc512](https://github.com/ChristopherVR/ooxml/commit/27cc5126cee1c351069e8eb7c2dafe4a0d7806a1))

### Refactor

- **ui:** Run pptx find and replace on one shared session ([#202](https://github.com/ChristopherVR/ooxml/issues/202)) ([61ff51e](https://github.com/ChristopherVR/ooxml/commit/61ff51e2710edad14212132b475652b27bb05d43))

## [1.27.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.27.0) - 2026-10-25

### Features

- **ui:** Describe the pptx text effects as shared field descriptors ([371ce97](https://github.com/ChristopherVR/ooxml/commit/371ce972368f74d2645500c41ba4ee581777fd55))
- **ui:** Share one print job plan across the powerpoint bindings ([e9e05ba](https://github.com/ChristopherVR/ooxml/commit/e9e05baf890a3dccbbd6df555e35944d7544430a))
- **ui:** Group the shared keyboard shortcut reference ([ab6c1fd](https://github.com/ChristopherVR/ooxml/commit/ab6c1fda0068e213b6753363fb26ed353d713c13))
- **pptx:** Let hosts disable or self-host the Google Fonts fallback ([#200](https://github.com/ChristopherVR/ooxml/issues/200)) ([e5de994](https://github.com/ChristopherVR/ooxml/commit/e5de994b1d976c64f727dbefd35ae640a4563ebf))

### Bug Fixes

- **ui:** Decide what a new slide takes once its layout arrives ([ce051bd](https://github.com/ChristopherVR/ooxml/commit/ce051bd9101627277eaa8602d911aa2b89ec2a30))
- **ui:** Share theme undo records and queue every theme part write ([981d069](https://github.com/ChristopherVR/ooxml/commit/981d069b2421e8d984a051225f532f8d06d55d48))
- **ui:** Share an in-place theme by path switch and its undo ([bcaaa76](https://github.com/ChristopherVR/ooxml/commit/bcaaa7635069c05a7b160a2f85fc90635a97f358))
- **ui:** Undo a theme switch without writing other masters' parts ([5fdf793](https://github.com/ChristopherVR/ooxml/commit/5fdf793d0625b064a289e05df285b8ee2f520058))
- **ui:** Draw a switched master's style references with its new part ([915a3e1](https://github.com/ChristopherVR/ooxml/commit/915a3e1034cc4608f68d315a2e4c89056548b94f))
- **ui:** Undo an Effects pick in the viewer and the saved theme ([b43768c](https://github.com/ChristopherVR/ooxml/commit/b43768ccfb742e6daa1fe143d825cb404f73cd20))
- **ui:** Name what a master command removed and the steps that reach it ([3fb7709](https://github.com/ChristopherVR/ooxml/commit/3fb77093c1285f41f26d844fc9057c1f64fc8af5))

### Refactor

- **ui:** Restore a crossed history theme in one shared call ([9fac9fd](https://github.com/ChristopherVR/ooxml/commit/9fac9fd48b58434883305d08d566ea1c26172d5c))
- **pptx:** Forget a deleted part's undo steps in one shared call ([1b5ddc4](https://github.com/ChristopherVR/ooxml/commit/1b5ddc4517a7b54ecf97baa56cb55f2e574eb802))
- **ui:** Take the text effect field fallbacks from the toggles ([49973d5](https://github.com/ChristopherVR/ooxml/commit/49973d55d48b08ab153c714a7077fef5f0a1f119))
- **ui:** Keep the shortcut group label map private ([1b669d9](https://github.com/ChristopherVR/ooxml/commit/1b669d90c9c7edc9b9dc3079ccdf459983190015))
- **pptx:** Remove react and vue modules nothing imports ([#199](https://github.com/ChristopherVR/ooxml/issues/199)) ([54359ae](https://github.com/ChristopherVR/ooxml/commit/54359ae107d7782c0df857ec31993f089b3775a6))

### Documentation

- **ui:** Say how a binding adopts a slide master view command ([6acbbf5](https://github.com/ChristopherVR/ooxml/commit/6acbbf58e540ab597ab07815a2855e6b2020b34e))

### Testing

- **pptx:** Pin a design fonts pick to the font scheme part ([5555df9](https://github.com/ChristopherVR/ooxml/commit/5555df91ffae18c8025c183d9f98aeefcc4093a2))

### Chores

- **pptx:** Drop the text effect strings nothing reads anymore ([5d28074](https://github.com/ChristopherVR/ooxml/commit/5d2807489dc5b463f38c40dc25ea3b5a8cac0a42))
- **pptx:** Drop two shortcut strings only Vue's old catalog read ([860622c](https://github.com/ChristopherVR/ooxml/commit/860622c74e0d2471a8c019c1c68d10b983e9278d))

## [1.26.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.26.2) - 2026-10-24

### Bug Fixes

- **pptx:** Fill the outer corners of a widened outline clip ([3775336](https://github.com/ChristopherVR/ooxml/commit/377533638d18237cd19b45620e9f2e2cd9317844))
- **pptx:** Match morph text the way powerpoint carries it ([76a2b88](https://github.com/ChristopherVR/ooxml/commit/76a2b8894e82a88e627e90c81003af9709927270))
- **pptx:** Keep outlines whole through morphs and on inset lines ([1440535](https://github.com/ChristopherVR/ooxml/commit/1440535af878fec7512ed72dfc386a49f2c798ee))
- **docx:** Paint page-view shapes without the chart frame's clip ([a07593b](https://github.com/ChristopherVR/ooxml/commit/a07593b5148ba870babdccd493dd8be826cd887a))
- **ui:** Let smartart outlines paint past the drawing's edge ([81f76ce](https://github.com/ChristopherVR/ooxml/commit/81f76ce3832197fa5686a5617dd37a6d09a3b123))
- **pptx:** Stop the fill showing along a widened outline's edge ([52251e3](https://github.com/ChristopherVR/ooxml/commit/52251e39b5e74b0674e19380eee93c963c983a54))

## [1.26.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.26.1) - 2026-10-23

### Bug Fixes

- **pptx:** Apply theme picks in place and keep theme tiles stable ([ecb6385](https://github.com/ChristopherVR/ooxml/commit/ecb6385a8e211b60461f842eb4a88945555796e5))

### Testing

- **ui:** Resolve ooxml-ui/chart-format-locales to source in vitest ([56e8522](https://github.com/ChristopherVR/ooxml/commit/56e8522f0386b6e8487471884ad47130bbacd45f))

## [1.26.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.26.0) - 2026-10-23

### Features

- **ui:** Add the shared chart data sheet element ([0a87377](https://github.com/ChristopherVR/ooxml/commit/0a873776e44afbd9f2c672bafdc42e941cfd3957))
- **ui:** Add headless pptx controllers, key dispatch and ai bridge ([e76abc1](https://github.com/ChristopherVR/ooxml/commit/e76abc1938f9f43cfd138230980bbd3002a042a6))
- **ui:** Add the ooxml-ui/bindings framework binding kit ([519313b](https://github.com/ChristopherVR/ooxml/commit/519313b1cdd6f2fd323e0f8023b38b116025cd1e))
- **ui:** Add shared SmartArt picker and pptx insert dialogs ([85085d9](https://github.com/ChristopherVR/ooxml/commit/85085d9f21e9dc9b54c08d36c8e63f63db0b3f88))
- **ui:** Share one chart Format pane across products ([07f92d9](https://github.com/ChristopherVR/ooxml/commit/07f92d98cc3726aade9c67c9b8c206a804339a0a))
- **xlsx:** Format charts in the shared chart Format pane ([f1b7e30](https://github.com/ChristopherVR/ooxml/commit/f1b7e309b76ee19828f3a83a4f8e5a43851c93a1))
- **ui:** Add the shared pptx autosave and recovery controllers ([318997a](https://github.com/ChristopherVR/ooxml/commit/318997a5984e5479dfc913d7cad7785458329fbb))
- **ui:** Add the shared pptx collaboration provider and session ([56cfc9d](https://github.com/ChristopherVR/ooxml/commit/56cfc9d1d0bb78e8856e76f8fc7265479f4d4814))
- **ui:** Draw the pptx effects inspector from shared field descriptors ([3e5afa8](https://github.com/ChristopherVR/ooxml/commit/3e5afa829164142fd4a97866db24efe861c6db1e))
- **ui:** Draw the pptx line options from shared field descriptors ([6194f50](https://github.com/ChristopherVR/ooxml/commit/6194f509e0715ce9d653a49742402693952cbfee))
- **ui:** Draw the pptx size and position fields from descriptors ([edd890a](https://github.com/ChristopherVR/ooxml/commit/edd890a291dc01af63cfdb8c245b131d298868bc))
- **ui:** Draw the pptx text box options from shared field descriptors ([bf34da5](https://github.com/ChristopherVR/ooxml/commit/bf34da5910f4623e189b0ff7b6ffe37f632dd949))
- **ui:** Insert and edit Word charts on the shared chart elements ([3227a15](https://github.com/ChristopherVR/ooxml/commit/3227a15a57ff0da6c6deb6bcd5511f734281a456))
- **chart:** Share the chart design galleries and add quick layouts ([c95c89a](https://github.com/ChristopherVR/ooxml/commit/c95c89aab3ac5cdd6cdb500e033d0b799533343e))
- **xlsx:** Add chart styles and quick layout to chart design ([a16785a](https://github.com/ChristopherVR/ooxml/commit/a16785aeb35dc33a4476c18589bb8b8287185879))
- **visio:** Edit data and format an inserted chart object ([edbd1eb](https://github.com/ChristopherVR/ooxml/commit/edbd1ebb49bffe107ea5b78c056d709879926ed9))
- **ui:** Add the shared smartart text pane element ([1005bdc](https://github.com/ChristopherVR/ooxml/commit/1005bdc9ab3236bae1fb72ca0e0f283f5deeffc8))
- **ui:** Edit and insert smartart in the word editor ([228f6b3](https://github.com/ChristopherVR/ooxml/commit/228f6b365b1343b53451670c4cc4a97f997ce179))
- **ui:** Count uncommitted inline text as an autosave edit ([555ea92](https://github.com/ChristopherVR/ooxml/commit/555ea92f0a0bc49f8a3d0f852aafda0758ca8276))
- **ui:** Resize Word charts with the shared drawing handle ([0716e69](https://github.com/ChristopherVR/ooxml/commit/0716e6974c133cd3d2851698612ae5b5248b3a0b))
- **ui:** Format Word charts in the shared chart Format pane ([43b01d2](https://github.com/ChristopherVR/ooxml/commit/43b01d219d56c5e1219277da519842da6a449ebb))
- **ui:** Edit and insert smartart in the excel editor ([2f44aae](https://github.com/ChristopherVR/ooxml/commit/2f44aae764c9d1ed0bc8ea8039b23964f8bc724b))
- **ui:** Read the Excel chart Format pane from ChartSpace ([7e7717c](https://github.com/ChristopherVR/ooxml/commit/7e7717ced6b37c0a5661f7b7ca13fc3f70a08dd8))
- **diagram:** Share the outline levels of a diagram's nodes ([d20e8cf](https://github.com/ChristopherVR/ooxml/commit/d20e8cf0bcfe2b9b8c5c6e06f5701adb0ef68673))
- **ui:** Add tools to the shared smartart text pane ([7056ec7](https://github.com/ChristopherVR/ooxml/commit/7056ec72e90c62603bc2fc6e85db07d64c3521a0))
- **ui:** Add the pptx-ui-smartart-text-pane alias ([89dc5f1](https://github.com/ChristopherVR/ooxml/commit/89dc5f16edb98cbab866132f7de9548562cd0d4b))
- **ui:** Add shape before and above to the smartart design ribbon ([6a1c8d8](https://github.com/ChristopherVR/ooxml/commit/6a1c8d88414c4d2673ea289da043032558c10db3))
- **ui:** Draw the pptx inspector Recent Colors on the shared colour grid ([#174](https://github.com/ChristopherVR/ooxml/issues/174)) ([8f99d4a](https://github.com/ChristopherVR/ooxml/commit/8f99d4a3b1bf0750b51279b374992ffe338fb33a))

### Bug Fixes

- **ui:** Keep the smartart picker css on design tokens ([a58dddd](https://github.com/ChristopherVR/ooxml/commit/a58dddd33cfb6a81ebc0100796d972a88e8bf756))
- **ui:** Make the chart Format pane host its own landmark ([9b247d8](https://github.com/ChristopherVR/ooxml/commit/9b247d8ee4886eef31141eacc7328eb726460285))
- **ui:** Close the smartart text pane when a new document opens ([5e7f7f4](https://github.com/ChristopherVR/ooxml/commit/5e7f7f420c101f078f6669829f21105b047dbc4a))
- **ui:** Use tokens for the chart data sheet lengths ([ba916fe](https://github.com/ChristopherVR/ooxml/commit/ba916fe276f76bc5bc6152e51fbd0b079dc377a0))
- **ui:** Keep an open inline draft outstanding on a clean editor ([032ca5e](https://github.com/ChristopherVR/ooxml/commit/032ca5e79f35715b644f5815cd901851ce338e63))
- **xlsx:** Never rewrite a chart part the model cannot carry ([ac136f7](https://github.com/ChristopherVR/ooxml/commit/ac136f7319080e129106b452504ca915310367f3))
- **ui:** Keep the smartart text pane attribute-free on construction ([ed54acf](https://github.com/ChristopherVR/ooxml/commit/ed54acfe2b3d7d6ddbebb909c0b7d31ad56aa60c))

### Other

- Shared chart data sheet (wave6-5) ([b546a73](https://github.com/ChristopherVR/ooxml/commit/b546a73c0eff334308473e8d970e72cd674534e0))
- Pptx headless controllers (wave6-2) ([41dc33b](https://github.com/ChristopherVR/ooxml/commit/41dc33be23af816e33e08983f29e08621d5ae32f))
- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))
- Shared smartart picker and insert dialogs (wave6-6) ([60c4372](https://github.com/ChristopherVR/ooxml/commit/60c437219dc3e495238bb0760e17a8036e511f41))
- Shared chart format pane (wave6-7) ([0324c80](https://github.com/ChristopherVR/ooxml/commit/0324c80633332e602ce813bc64b0e231ab6b64af))
- Pptx autosave, recovery and collaboration controllers (wave6-10) ([e48da7a](https://github.com/ChristopherVR/ooxml/commit/e48da7a6ffb025682d27314d537868355512916f))
- Pptx inspector field descriptors (wave6-12) ([a360ffc](https://github.com/ChristopherVR/ooxml/commit/a360ffcbec5303ebe7d427c5e442f99d82e72d06))
- Word chart editing (wave6-11) ([45ecd4b](https://github.com/ChristopherVR/ooxml/commit/45ecd4b144f4195cfa4cbe85763b87fbca162470))
- Shared chart design galleries and Visio chart editing (wave6-14) ([692e476](https://github.com/ChristopherVR/ooxml/commit/692e4763f28083d3d280b78a2e1a73e13f77499b))
- Bring wave6-integration into wave6-13 ([0b881ed](https://github.com/ChristopherVR/ooxml/commit/0b881ed1a85fe81295e0e68f7f92bd137e65f91b))
- Word SmartArt editing (wave6-13) ([751ba15](https://github.com/ChristopherVR/ooxml/commit/751ba157ca578e8ff6b57371a243180101ed562e))
- Autosave uncommitted inline text (wave6-10b) ([d75cc4d](https://github.com/ChristopherVR/ooxml/commit/d75cc4dc43d86717653481c6237679117ad586a6))
- Word chart Format pane, chart resizing and shared pane translations (wave6-17) ([d6ac348](https://github.com/ChristopherVR/ooxml/commit/d6ac3489fd3e8ed959f4978b4de8fd4674ee71b3))
- Shared pptx canvas gestures, group selection and show steps (wave6-16) ([2f0df81](https://github.com/ChristopherVR/ooxml/commit/2f0df81355f1c3e4db3e12a8d39489b5d4c23825))
- Excel SmartArt editing on the shared diagram core (wave6-20) ([c5fdb99](https://github.com/ChristopherVR/ooxml/commit/c5fdb9919cfce94ae9daac4d1079ca235c0c5183))
- Excel chart editing over ChartSpace (wave6-18) ([ce75f7f](https://github.com/ChristopherVR/ooxml/commit/ce75f7faba40314ebc8c6de2a43c757328aa4c40))
- Pptx SmartArt text pane on the shared element (wave6-19) ([e025597](https://github.com/ChristopherVR/ooxml/commit/e025597647c9b19da2c8a308e0cc44734c569c4d))
- Bring origin/main into wave 6 ([2fedfaa](https://github.com/ChristopherVR/ooxml/commit/2fedfaae708c279044fdf86ea192315e366f80da))
- Ooxml-ui and shared element docs (docs6-1) ([c1635a1](https://github.com/ChristopherVR/ooxml/commit/c1635a146f796af9f810eebc2befb81546e40ea9))

### Refactor

- **teams:** Move bindTeams into ooxml-ui/teams on the binding kit ([fd417a5](https://github.com/ChristopherVR/ooxml/commit/fd417a515f76b580af69acc5b1e319d15d41fee5))
- **diagram:** Move built-in smartart definitions to diagram/builtin ([d673145](https://github.com/ChristopherVR/ooxml/commit/d6731456c840ef85032e490f60c3492b36cd7445))
- **ui:** Ship the chart Format pane's own translations ([d6c3d65](https://github.com/ChristopherVR/ooxml/commit/d6c3d6539cf8f28ba5032959a7e2bbd0076f15ae))
- **ui:** Run every pptx canvas gesture on the shared controller ([f10083a](https://github.com/ChristopherVR/ooxml/commit/f10083abf1c3dfc34dd7d2b0e0107ad8dadd0343))
- **ui:** Decide pptx slide show steps once for every binding ([582e31f](https://github.com/ChristopherVR/ooxml/commit/582e31ff532e343f589e3345360783e275704172))
- **diagram:** Share the smartart ui strings with every product ([b96c393](https://github.com/ChristopherVR/ooxml/commit/b96c393e820e88317bda0dea1ab25eb9e20b8b26))
- **ui:** Share the smartart text pane wiring and design list ([5a051b5](https://github.com/ChristopherVR/ooxml/commit/5a051b50dc9983f8979d3bdf96224254fd901fa0))

### Documentation

- **ui:** Document the ooxml-ui/bindings kit and its provenance ([7f53c90](https://github.com/ChristopherVR/ooxml/commit/7f53c90123674e26962d5a0f6e94718c4f1314e1))
- **pptx:** Record the autosave, recovery and collaboration controllers ([6bbbef7](https://github.com/ChristopherVR/ooxml/commit/6bbbef778175cc320c152f4587014dc435d6d8b6))
- **ui:** Document the binding kit, chart and SmartArt elements ([a3dd7b8](https://github.com/ChristopherVR/ooxml/commit/a3dd7b81d6ad1fe639e8f8c929556ea4c1205289))

### Chores

- **ui:** Regenerate the custom elements manifest after the merges ([ebb0650](https://github.com/ChristopherVR/ooxml/commit/ebb06500d5fe9d43bfee8a9773bae6a01307bd4c))

## [1.25.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.25.0) - 2026-10-23

### Features

- **ui:** Add a shared progress dialog for the pptx export modal ([#175](https://github.com/ChristopherVR/ooxml/issues/175)) ([05ca9c0](https://github.com/ChristopherVR/ooxml/commit/05ca9c0a9c3fd4726fed6ae9cacd9736512d30a3))
- **ui:** Share the pptx read-only banner session ([#176](https://github.com/ChristopherVR/ooxml/issues/176)) ([43effb0](https://github.com/ChristopherVR/ooxml/commit/43effb07fa8a71977557b345da1913de3646a59a))

### Bug Fixes

- **ui:** Keep a selected float's handles through a print relayout ([#177](https://github.com/ChristopherVR/ooxml/issues/177)) ([eeda2bf](https://github.com/ChristopherVR/ooxml/commit/eeda2bfe52a643ffa9c878bbad15502e7b14e059))

### Performance

- **ui:** Let html2canvas clone only the captured element's branch ([#171](https://github.com/ChristopherVR/ooxml/issues/171)) ([ee0cd6e](https://github.com/ChristopherVR/ooxml/commit/ee0cd6e8045c2f0271c511909981dbe7d10677be))

## [1.24.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.24.0) - 2026-10-22

### Features

- **ui:** Move the pptx inspector Quick Styles onto the shared gallery ([#168](https://github.com/ChristopherVR/ooxml/issues/168)) ([978fdbd](https://github.com/ChristopherVR/ooxml/commit/978fdbdb82335f24a0496ba1bf1152a7018cee63))
- **ui:** Build the pptx status bar state in one place ([#169](https://github.com/ChristopherVR/ooxml/issues/169)) ([0af0608](https://github.com/ChristopherVR/ooxml/commit/0af06085a54134ef3679d66f3319c50d8527a960))

### Bug Fixes

- **docx:** Follow-up fixes for content controls in table cells ([#166](https://github.com/ChristopherVR/ooxml/issues/166)) ([47b0e0f](https://github.com/ChristopherVR/ooxml/commit/47b0e0ffde6a022b40969936acb9c734de4cb8cc))
- **docx:** Keep a protected form's text field values unformatted ([#167](https://github.com/ChristopherVR/ooxml/issues/167)) ([cb21dff](https://github.com/ChristopherVR/ooxml/commit/cb21dff2988184e192b30192b807ff037a1e894e))
- **ui:** Keep 3d smartart fill swatches while crossing a node label ([1099185](https://github.com/ChristopherVR/ooxml/commit/1099185aa0c1b7f1a0b4600ec0863ba879e2be40))
- **pptx:** Keep the outer half of centred outlines on clipped shapes ([df3ca02](https://github.com/ChristopherVR/ooxml/commit/df3ca028919989f55a8a5fdc3c3cb8996c09987e))
- **ui:** Paint the morph overlay in one stacking order ([9f5a11f](https://github.com/ChristopherVR/ooxml/commit/9f5a11fa15c3c31eb06551decda93b96b7841b94))

### Testing

- **pptx:** Cover the stroked clip widening in every binding ([0c29435](https://github.com/ChristopherVR/ooxml/commit/0c2943500ac3026066aa699d26ed758bca2891bd))

## [1.23.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.23.1) - 2026-10-22

### Bug Fixes

- **ui:** Fit the pptx slide to sub-pixel viewport sizes ([#165](https://github.com/ChristopherVR/ooxml/issues/165)) ([d09a71d](https://github.com/ChristopherVR/ooxml/commit/d09a71d1e478eebc28edec2af309e7f4b0a69dbb))

### Testing

- **ci:** Check packed browser entries for node-only modules ([0420bef](https://github.com/ChristopherVR/ooxml/commit/0420bef6f98e8718a9051bc4a945980a57edf435))

## [1.23.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.23.0) - 2026-10-22

### Features

- **visio:** Add file-command and host-owned storage to the viewer ([0c8e284](https://github.com/ChristopherVR/ooxml/commit/0c8e284015fe1cc1b895c840900482ba12f8cda4))
- **docx:** Route save and open through host-owned storage ([ec27cd2](https://github.com/ChristopherVR/ooxml/commit/ec27cd28c4f594446e08b60ab05db2b64b72c683))
- **xlsx:** Route save and open through host-owned storage ([ef768a6](https://github.com/ChristopherVR/ooxml/commit/ef768a684f558275071be347d68406503aef8c23))
- **pptx:** Route saves and recovery through host-owned storage ([4ad5655](https://github.com/ChristopherVR/ooxml/commit/4ad565525269ef0284114efad6d5e4a7bed756e4))
- **ui:** Host regions, file tab pages and menu commands for every editor ([f2b24f1](https://github.com/ChristopherVR/ooxml/commit/f2b24f1ddb929b299a1a8b5d6fa887ca45921941))
- **docx:** Take the shared customization contract in <docx-editor> ([65f05d3](https://github.com/ChristopherVR/ooxml/commit/65f05d3d70cc28e6633f300fb4758aca7ba34184))
- **visio:** Take the shared customization contract in <visio-viewer> ([3a6bb30](https://github.com/ChristopherVR/ooxml/commit/3a6bb305b8ac04f0942b40e8cdfbd887093c6112))
- **xlsx:** Take the shared customization contract in <xlsx-editor> ([f96a544](https://github.com/ChristopherVR/ooxml/commit/f96a5449a6a36fba59997197b484ddefb4f2d8e8))
- **embedding:** Tidy menu separators around hidden entries ([c4e6101](https://github.com/ChristopherVR/ooxml/commit/c4e6101b620e8ef9bc1c71387203232833015cf0))
- **xlsx:** Offer host menu commands on header and sheet-tab menus ([ebef1b8](https://github.com/ChristopherVR/ooxml/commit/ebef1b8a1102c568b1a7a89ab5c3390f7887aa82))
- **visio:** Accept readOnly as the inverse of canEdit ([fbc325b](https://github.com/ChristopherVR/ooxml/commit/fbc325be8c9c12a7c2e4a3affd7af2428ba4ed49))
- **pptx:** Add readOnly to the five viewer bindings ([4aadd67](https://github.com/ChristopherVR/ooxml/commit/4aadd678969d4b2d4811d86880b91c79ba7ca3b6))
- **pptx:** Add a shared autosave status kind for host events ([4b1f7fb](https://github.com/ChristopherVR/ooxml/commit/4b1f7fb0ed84b5ab309a155039a1f0a23f197bb2))
- **pptx:** Report the missing host events from the react viewer ([2ba8c58](https://github.com/ChristopherVR/ooxml/commit/2ba8c58d45446cf912eab5d3c4c026a55fd50a3e))
- **pptx:** Report the missing host events from the vue viewer ([37167d2](https://github.com/ChristopherVR/ooxml/commit/37167d234b435780fcfe63bd4733a413ff84f11a))
- **pptx:** Report the missing host events from the svelte viewer ([08bdc7d](https://github.com/ChristopherVR/ooxml/commit/08bdc7dc04ef235c86670df8747022f3718f38cb))
- **pptx:** Report the missing host events from the angular viewer ([94b9de4](https://github.com/ChristopherVR/ooxml/commit/94b9de473167c68e3b678636f6453af237d7029c))
- **pptx:** Report the missing host events from the vanilla viewer ([d6590fd](https://github.com/ChristopherVR/ooxml/commit/d6590fd6f61ca77f92a6b49001445eb434c548b0))
- **embedding:** Land the host customization, read-only and storage contracts ([822a6a0](https://github.com/ChristopherVR/ooxml/commit/822a6a073840b9e9d1671476542322b28332dc42))

### Bug Fixes

- **pptx:** Scope the viewer stylesheet reset to the viewer roots ([c390b2b](https://github.com/ChristopherVR/ooxml/commit/c390b2ba2f0ec7811c3554b9989346c6e8cd927d))
- **visio:** Keep autorecover until a host save succeeds ([02dab92](https://github.com/ChristopherVR/ooxml/commit/02dab92d60fed753538b1363baa5ece11488577f))
- **docx:** Keep edits made during a host save unsaved ([019bc7b](https://github.com/ChristopherVR/ooxml/commit/019bc7b33bade1dd16565507866d309be2479684))
- **xlsx:** Keep edits made during a host save unsaved ([03f2a86](https://github.com/ChristopherVR/ooxml/commit/03f2a863bc262fad1fab8a3fc5713702f769bf95))
- **pptx:** Report host save failures and use host recovery everywhere ([8e982c8](https://github.com/ChristopherVR/ooxml/commit/8e982c8576f656d10200a45cc1708a5ac8e14ed6))
- **xlsx:** Hide context-menu entries whose action is hidden ([f28dbf8](https://github.com/ChristopherVR/ooxml/commit/f28dbf85271455c1151587d475794b4cf167c2e6))
- **visio:** Hide Quick Access Save and tidy menus under capabilities ([f33e6b9](https://github.com/ChristopherVR/ooxml/commit/f33e6b9d766de0ae995d651714b1fb5fe2652059))
- **docx:** Keep capability-hidden commands out of Customize Ribbon ([845285f](https://github.com/ChristopherVR/ooxml/commit/845285f35873cb6f72ddcd662227f0644b877c27))
- **xlsx:** Keep a viewer role read-only when the host clears readOnly ([fe3a56b](https://github.com/ChristopherVR/ooxml/commit/fe3a56be1d0100543bb170627b99a966d4912b52))
- **docx:** Make a viewer collaboration role go through readOnly ([d2fd1a8](https://github.com/ChristopherVR/ooxml/commit/d2fd1a8dd63fc5a760ebf45b5203d569287e0d27))
- **docx:** Report read-only with the lock and host switch ([735fb0d](https://github.com/ChristopherVR/ooxml/commit/735fb0d78ff18dc71eeeb83db2a76f26e423ba30))
- **xlsx:** Report read-only with the lock and host switch ([782adb9](https://github.com/ChristopherVR/ooxml/commit/782adb95a4c8ad4dded13d73e8a5fa5aee2c8666))
- **xlsx:** Keep context menu groups apart when an entry is hidden ([8364768](https://github.com/ChristopherVR/ooxml/commit/8364768bb879abda5cb7b14462ab956012df6fa4))
- **docx:** Honour find: false in Ctrl+F and the shortcut help ([f38d2fa](https://github.com/ChristopherVR/ooxml/commit/f38d2fae33cad0444e201ab99635d56a985c9b26))
- **xlsx:** Use the shared menu separator rule and test the locks ([9f07a35](https://github.com/ChristopherVR/ooxml/commit/9f07a35cd83bd6ae64bc9300f9c7220879b29a91))
- **visio:** Keep capability-hidden commands out of Tell me ([c5d6bb9](https://github.com/ChristopherVR/ooxml/commit/c5d6bb99ca9d21adba12f30e5546bccd9cf10fff))

### Other

- Integrate the customization contract into host storage ([629554b](https://github.com/ChristopherVR/ooxml/commit/629554b3030728bfe16e11ac8ae5b2530bb28a17))
- Integrate the pptx binding events into host storage ([08272a2](https://github.com/ChristopherVR/ooxml/commit/08272a298c8b249aa0bf143b689ff01f66663309))
- Bring origin/main into the embedding work ([a3711a1](https://github.com/ChristopherVR/ooxml/commit/a3711a14049a2cd11b994e76eafec039db2dc2de))

### Performance

- **visio:** Keep the parse worker alive between page reads ([50638ff](https://github.com/ChristopherVR/ooxml/commit/50638ffcdc4789c390147d6d32b7365345e47964))
- **ui:** Re-export pptx text helpers from core by name ([b120642](https://github.com/ChristopherVR/ooxml/commit/b12064294a367faf44da377c55d1255b2fa75729))

### Refactor

- **pptx:** Base the customization types on the shared contract ([60b71d3](https://github.com/ChristopherVR/ooxml/commit/60b71d3727de83a3c25c6bb7e5badba70d1037a6))

## [Unreleased]

### Documentation

- Fix stale paths, counts and provenance in area docs ([95ad5db](https://github.com/ChristopherVR/ooxml/commit/95ad5db34b977df09e9a563b6ac72305cddd8ae7))
- Point script, e2e and plan references at real paths ([eec05e2](https://github.com/ChristopherVR/ooxml/commit/eec05e293e6149c0ecf2dd3c863694721949294a))
- **ui:** Describe the ui decisions page as it is now ([c37c73a](https://github.com/ChristopherVR/ooxml/commit/c37c73afb4c4bf6bce8b637c524d76184b44689c))

## [1.22.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.22.0) - 2026-10-22

### Features

- **xlsx:** Let hosts own the data-validation alert in building blocks ([fcbe367](https://github.com/ChristopherVR/ooxml/commit/fcbe367f5cfe366cc425796e3fdff9c31aa99e25))
- **visio:** Expose the text and connector tools on the standalone canvas ([8b5b094](https://github.com/ChristopherVR/ooxml/commit/8b5b09450f875c000415cee721fe1a6c74bd8ccf))
- **visio:** Confirm owner write-back and report its status ([ef58158](https://github.com/ChristopherVR/ooxml/commit/ef5815875434f3ce62ef0f743a16c568e8a25e45))
- **pptx:** Confirm owner write-back and report its status ([2cb779f](https://github.com/ChristopherVR/ooxml/commit/2cb779f606c7d8d0778c11be830433fc93f728c3))

## [1.21.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.21.0) - 2026-10-22

### Features

- **pptx:** Report the shared collaboration state from every binding ([62752ef](https://github.com/ChristopherVR/ooxml/commit/62752efd7b2f9df8372db485cd23c2a6ed5f15ee))
- **ui:** Complete the docx surface building blocks ([e3b34f9](https://github.com/ChristopherVR/ooxml/commit/e3b34f96dcaacd740ab7f7272b5567a6e287f392))
- **xlsx:** Report validation, host dialogs and state in surface props ([11ce627](https://github.com/ChristopherVR/ooxml/commit/11ce6271694e47c6fccd3d40040a44d1ae8ca4c9))
- **ui:** Give the standalone visio canvas the full viewer's editing ([672522f](https://github.com/ChristopherVR/ooxml/commit/672522ffd012adf4bf6526aadc5c3bfdfbcca95e))
- **pptx:** Connect through a provider, webrtc or same-browser room ([4ae5288](https://github.com/ChristopherVR/ooxml/commit/4ae528882a33ec93730f30bf642589a00b6edf06))
- **visio:** Merge collaboration per shape, with cursors and write-back ([3d1564c](https://github.com/ChristopherVR/ooxml/commit/3d1564caa98d86b97bd7a384bd9d11fe533435f7))

### Bug Fixes

- **pptx:** Cover transition direction names the travel, not the edge ([e9a53d3](https://github.com/ChristopherVR/ooxml/commit/e9a53d3ae5fe409b035b4c37929605eccad2e373))
- **pptx:** Do not morph-pair pictures that paint different images ([87f9a03](https://github.com/ChristopherVR/ooxml/commit/87f9a03d162803de7189e812a0455e5ac3fb7398))
- **pptx:** Morph a picture's transparency change instead of popping it ([16485be](https://github.com/ChristopherVR/ooxml/commit/16485be1649e09a4c5d81f12fddc401805dd6389))
- **pptx:** Do not morph-pair pictures that paint different images ([a60887d](https://github.com/ChristopherVR/ooxml/commit/a60887d1bda1af1d067a0adfa2b9673a6eae62d3))
- **xlsx:** Name the share opener like the other building blocks ([1346849](https://github.com/ChristopherVR/ooxml/commit/1346849e085724ef8a4810cf540f8d802fd47cee))
- **docx:** Name the toolbar openers apart from the host options ([d5bc691](https://github.com/ChristopherVR/ooxml/commit/d5bc691e2aaa946147e55edf590a61f64c7a671a))

## [1.20.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.20.0) - 2026-10-21

### Features

- **ui:** Draw and operate content controls around table cells ([70dcd1e](https://github.com/ChristopherVR/ooxml/commit/70dcd1ec8550b8e469ea8f9c94c6104241e7c90f))
- **docx:** Fill legacy form fields and honour forms protection ([8ab5a99](https://github.com/ChristopherVR/ooxml/commit/8ab5a99954432d51ddf8df92c4d6a7dba619de8e))
- **docx:** Show the STYLEREF text of each page in Print Layout headers ([b82021a](https://github.com/ChristopherVR/ooxml/commit/b82021a062724fcdf96a45e88423f21b9382f380))
- **docx:** Update body STYLEREF fields with Update Fields ([985b721](https://github.com/ChristopherVR/ooxml/commit/985b7216deea2cfc274a9294d578c696d39733ee))
- **ui:** Add shared Help and Record ribbon models for pptx ([415cb6a](https://github.com/ChristopherVR/ooxml/commit/415cb6a2acbad4655daf910e3a78284215f91039))
- **ui:** Let the colour grid take a product's palette and selection ([81928c1](https://github.com/ChristopherVR/ooxml/commit/81928c17af5f4a6c00d1f2aac8da8514e051ae9f))
- **ui:** Let the colour grid take a product's palette and selection ([9993199](https://github.com/ChristopherVR/ooxml/commit/99931990b8db4fe7c495d9d7f771e7087d23968f))
- **ui:** Add a pptx colour grid alias and a deck palette adapter ([a98cb8b](https://github.com/ChristopherVR/ooxml/commit/a98cb8b8a1ab2a4c6dc78b5c7f7d25f519bc6b9a))
- **ui:** Let the colour grid fit a narrow pane ([8313b65](https://github.com/ChristopherVR/ooxml/commit/8313b65e5c53caab9adaeefa4a39dad1c77930e0))
- **ui:** Add a compact gallery panel for panes that apply at once ([20a95ad](https://github.com/ChristopherVR/ooxml/commit/20a95ad80844efe617d2e6cdbeffc0d2ba615b2d))
- **ui:** Add the pptx Text Warp gallery ([3dfd6da](https://github.com/ChristopherVR/ooxml/commit/3dfd6da2c813a87ee984aaef92ccdeba36f184d5))

### Bug Fixes

- **ui:** Refuse an incompatible pptx control before writing the theme ([da3b217](https://github.com/ChristopherVR/ooxml/commit/da3b217b59082268bd58dcd964af88b5e1854222))
- **ui:** Hide the xlsx ribbon overflow button once its group unfolds ([c6240f0](https://github.com/ChristopherVR/ooxml/commit/c6240f051000c6ea2f6f36f2d1de7560887c333c))
- **ui:** Keep pptx status bar touch targets at 44px ([31b63f9](https://github.com/ChristopherVR/ooxml/commit/31b63f9f8bba36b9c0721c70506789dd132fcdcb))
- **ui:** Keep a table cell's content control when the cell is copied ([0bfa4ae](https://github.com/ChristopherVR/ooxml/commit/0bfa4ae5d7a76b8f86654f6de237e96895f1244e))
- **ui:** Operate the content control that was clicked ([7350c8e](https://github.com/ChristopherVR/ooxml/commit/7350c8efbd22a6cd77e7b25330fad1e75f497c75))
- **ui:** Save the row controls of a merged table as the editor has them ([bf09c28](https://github.com/ChristopherVR/ooxml/commit/bf09c28e152ac9d52f551ada674a9ceaafec2ada))
- **docx:** Write a new control's verbatim properties only as properties ([a5db4d4](https://github.com/ChristopherVR/ooxml/commit/a5db4d452204fe08f993e69e4fb4104bdd4aeea6))
- **docx:** Remove a temporary block control that holds a table ([5270cfa](https://github.com/ChristopherVR/ooxml/commit/5270cfa67fb8b0935588e92e22ba8b298c7549c3))
- **docx:** Take a content control from pasted HTML only when it is one ([72e8a53](https://github.com/ChristopherVR/ooxml/commit/72e8a5310895e2ea5022921ec768b7c718adc6fa))
- **docx:** Check pasted cell, row and block controls like inline ones ([5556931](https://github.com/ChristopherVR/ooxml/commit/5556931fcfc309ab3811e21e23468fa1b01d823d))
- **docx:** Keep a cell control the editor holds none for ([4c5d208](https://github.com/ChristopherVR/ooxml/commit/4c5d2085f0429ee3158e57cdc09aeb6b682ba2da))
- **docx:** Apply a cell control's rules to a paste ([4a59e85](https://github.com/ChristopherVR/ooxml/commit/4a59e85b592c21601e57b9ee28e537fd1b108245))
- **docx:** Take and write only values of a control's own kinds ([cd0d434](https://github.com/ChristopherVR/ooxml/commit/cd0d434b0c03f4b691dd412d2ac9fdaf6faddbe1))
- **docx:** Take no pasted control nested deeper than a control is ([0578920](https://github.com/ChristopherVR/ooxml/commit/057892030a4a39478d50cba97fa1f528bb5857d9))
- **docx:** Edit paragraphs that hold proofing marks, symbols and smart tags ([8d2b489](https://github.com/ChristopherVR/ooxml/commit/8d2b48976c96322b0be46a5ee1a716d8f500fcb3))
- **docx:** Honour content control locks around legacy form fields ([1a7e72f](https://github.com/ChristopherVR/ooxml/commit/1a7e72fd20b89022fc197ecce7cd4eab07c99ba2))
- **docx:** Keep headers, footers and notes read-only under forms protection ([4be8e69](https://github.com/ChristopherVR/ooxml/commit/4be8e69364180135598db377a0e781fe7182909c))
- **docx:** Place unprotected sections of a protected form as layout does ([23e746c](https://github.com/ChristopherVR/ooxml/commit/23e746c857a46baf3ac2a11f9a610900c2a9212a))
- **docx:** Let the editor renumber split paragraphs in a protected form ([3e338e5](https://github.com/ChristopherVR/ooxml/commit/3e338e5017a620acad9b0b011e09c5ec91f5dcd3))
- **docx:** Operate form check boxes and drop-downs from the keyboard ([b0eb45d](https://github.com/ChristopherVR/ooxml/commit/b0eb45dd8f780b0dd82b3d11fc2e5a309c7d6ff9))
- **docx:** Operate form fields only at the caret the browser shows ([e20eb78](https://github.com/ChristopherVR/ooxml/commit/e20eb78776152161855c90746760720e5315de5b))
- **docx:** Keep markers at their characters in merged tracked runs ([ddab61b](https://github.com/ChristopherVR/ooxml/commit/ddab61b8e7da6bddb6e5048fd4fccdbe0a986dfd))
- **docx:** Keep a hyperlink's markup when its whole label is replaced ([57e059d](https://github.com/ChristopherVR/ooxml/commit/57e059d7b3fb4b3515277c81c0a83b4388dc4bc8))
- **docx:** Find words that hold soft hyphens ([7dc06aa](https://github.com/ChristopherVR/ooxml/commit/7dc06aa7e6771e95779a9752e63720f1b0144fd0))
- **docx:** Change run boundaries over identical unmodeled properties ([38c2ad8](https://github.com/ChristopherVR/ooxml/commit/38c2ad8e45f833c800f430c9b6ac5e58cd25ab1a))
- **docx:** Pick the drop-down entry chosen when names repeat ([02389cc](https://github.com/ChristopherVR/ooxml/commit/02389cc4c61e77a918426f84b51748359f5f495b))
- **docx:** Pick in the drop-down field the list was opened for ([b1625b6](https://github.com/ChristopherVR/ooxml/commit/b1625b6909f3d9945ed4efef35d656f7d2d87281))
- **docx:** Keep a form field's bookmark around it when the field moves ([42dcb86](https://github.com/ChristopherVR/ooxml/commit/42dcb86ba0f9a2668c22f8463cc02ab4272df847))
- **docx:** Keep a pasted tab ordinary beside a positional tab ([e31ea3a](https://github.com/ChristopherVR/ooxml/commit/e31ea3a1f8fef35058c6dd239bac8a88f0555cb7))
- **docx:** Keep a protected form's fields from taking formatting ([e137fb6](https://github.com/ChristopherVR/ooxml/commit/e137fb69fbfee30f6e7661d34208deeb85078cfb))
- **docx:** Format a CHARFORMAT STYLEREF result like its field name ([8030a2f](https://github.com/ChristopherVR/ooxml/commit/8030a2f04c5e10aa1c9d37ffae987d1796d676fb))
- **docx:** Count a paragraph box without lines as shown on its page ([3b91863](https://github.com/ChristopherVR/ooxml/commit/3b918630519025cbdf75a136fb672e94eb1bdfc2))
- **docx:** Clear a body STYLEREF result in Update Fields ([b173820](https://github.com/ChristopherVR/ooxml/commit/b1738206c0606a95963611fad547d8eb466560c8))
- **docx:** Reformat a CHARFORMAT STYLEREF result whose text is current ([c5687d5](https://github.com/ChristopherVR/ooxml/commit/c5687d5bcc7afdae91d3f1f721e2b22cfa08ef98))
- **docx:** Keep line breaks in a field result when updating fields ([b6de6e4](https://github.com/ChristopherVR/ooxml/commit/b6de6e45548af97772a4cff70b49f99765794322))
- **docx:** Keep field results out of text a revision deletes ([c6b59e4](https://github.com/ChristopherVR/ooxml/commit/c6b59e4d16aa87f65e1ea7b5716d162e01e5143d))
- **docx:** Update fields that read other fields in one Update Fields ([4700386](https://github.com/ChristopherVR/ooxml/commit/47003860f98f36abfdd2d019618d8c06c98e2a4c))
- **pptx:** Render weight-named font families at their weight ([7188288](https://github.com/ChristopherVR/ooxml/commit/718828829585145b461742c04e030fe3fbbd2d52))
- **pptx:** Render style-named font families in their style ([5413be6](https://github.com/ChristopherVR/ooxml/commit/5413be6b74a6409c8a2f1454e1a1cf4c4f1b8ba8))
- **pptx:** Apply named weights and styles to Vue WordArt and bullet fonts ([fe6e687](https://github.com/ChristopherVR/ooxml/commit/fe6e687b27c70831b42b4e5d152cbfe5292774c1))
- **ui:** Stretch the ribbon section to fill the ribbon row ([09a3996](https://github.com/ChristopherVR/ooxml/commit/09a399638121bdf93f821ef6ac71e31b46e5c0dc))
- **ui:** Size the colour grid's theme rows to its palette ([cf54101](https://github.com/ChristopherVR/ooxml/commit/cf541015d1015edae897d37298c7c5f10c89cdff))
- **color:** Require each palette variant to name its step ([70d7d9a](https://github.com/ChristopherVR/ooxml/commit/70d7d9a949e9dc5289bacf33bf63bd9af4c9f258))
- **ui:** Keep a short palette column in place in the colour grid ([f7b8724](https://github.com/ChristopherVR/ooxml/commit/f7b872425606ba592699a796dc87123d8c8efa38))
- **ui:** Size the colour grid's theme rows to its palette ([b9b56b5](https://github.com/ChristopherVR/ooxml/commit/b9b56b506d9948d3718d51a65fff874cebfd05b4))
- **color:** Require each palette variant to name its step ([3135862](https://github.com/ChristopherVR/ooxml/commit/3135862978fd352498d977335777555ba519428a))
- **ui:** Keep a short palette column in place in the colour grid ([8297a4b](https://github.com/ChristopherVR/ooxml/commit/8297a4b1e694670c9f98b4f74d544e89800b4185))
- **ui:** Let the pptx ribbon gallery draw as a panel ([b016984](https://github.com/ChristopherVR/ooxml/commit/b0169849e9f1496c660d34d919b8705b98c81686))
- **docx:** Write a new content control's properties only when valid ([3a8e23a](https://github.com/ChristopherVR/ooxml/commit/3a8e23ac2df41db82357568749240592eab361a5))
- **docx:** Keep foreign verbatim properties out of pasted controls ([9ee6664](https://github.com/ChristopherVR/ooxml/commit/9ee666419ca8fdcf4bea3a3f58753011057a798f))
- **docx:** Keep a copied control's binding and check extension types ([19e6e78](https://github.com/ChristopherVR/ooxml/commit/19e6e785a34c94e0e8bf86e92dc4def22e0972fc))
- **docx:** Take and write only values of a control's own kinds ([cb4b8e1](https://github.com/ChristopherVR/ooxml/commit/cb4b8e148a3666da6abf9175cc3760699a38a4a0))
- **pptx:** Cover transition direction names the travel, not the edge ([1b8e853](https://github.com/ChristopherVR/ooxml/commit/1b8e853abd84290590fd8402252bd1b47791af62))
- **pptx:** Morph a picture's transparency change instead of popping it ([cf1defe](https://github.com/ChristopherVR/ooxml/commit/cf1defe0c0a6fe9f41e46496fd1801edf028484b))

### Testing

- **ui:** Keep row content controls when text is typed in their table ([714de26](https://github.com/ChristopherVR/ooxml/commit/714de26613dc4a37985e6fbe4aa23e2920f63e55))
- **ui:** Keep block controls around a table and in a cell on save ([0702ce5](https://github.com/ChristopherVR/ooxml/commit/0702ce5ed69845836fe57841fbbbd2774da5b810))
- **ui:** Drop dragged content controls through the editor's handler ([dfdf11c](https://github.com/ChristopherVR/ooxml/commit/dfdf11c5324c4a89bbe6cdb81b34572bdd1eb071))

## [1.19.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.19.0) - 2026-10-21

### Features

- **docx:** Expose composable editor building blocks ([79bf480](https://github.com/ChristopherVR/ooxml/commit/79bf4801d8efb256a3e3e6fa81cca2613e6fd0fd))
- **xlsx:** Expose composable spreadsheet building blocks ([9aef499](https://github.com/ChristopherVR/ooxml/commit/9aef499f76196be6dd008ae873ac8f0d4caeeb58))
- **visio:** Expose canvas building blocks and host permissions ([e464a1c](https://github.com/ChristopherVR/ooxml/commit/e464a1c0399434430925e0e59b430b5788f5d8d9))
- **ui:** Share scoped themes across native building blocks ([fdd349c](https://github.com/ChristopherVR/ooxml/commit/fdd349c470866b73c22722c9967052fd41e851c4))
- **visio:** Theme native building blocks with shared office tokens ([f86b882](https://github.com/ChristopherVR/ooxml/commit/f86b882e2081540997f14af754805c573b372dc9))
- **docx:** Inherit shared office themes across editor blocks ([3adb658](https://github.com/ChristopherVR/ooxml/commit/3adb658f212e93a49beec631a618c8802b0c183a))
- **xlsx:** Share office themes across editor building blocks ([a791ec2](https://github.com/ChristopherVR/ooxml/commit/a791ec2e68956921006c719397803e603286a8c0))

### Bug Fixes

- **pptx:** Avoid morphing unrelated picture artwork ([294db0b](https://github.com/ChristopherVR/ooxml/commit/294db0b90b1a667af15ab6b4beb30901461b5dd4))
- **ui:** Harden composition lifecycle and consumer checks ([6aa52be](https://github.com/ChristopherVR/ooxml/commit/6aa52be88175ca4f46f6f3b1c223ef06a8f1f915))
- **ui:** Preserve live formula drafts and runtime compatibility ([3a7d4f1](https://github.com/ChristopherVR/ooxml/commit/3a7d4f184658c4d69b04bbd36bab7afa3896dd9e))
- **ui:** Retain scoped fonts on standalone native fields ([8e1a824](https://github.com/ChristopherVR/ooxml/commit/8e1a82401db83b3d8e85749ba1cfe9a1a8b69ab1))
- **ui:** Install scoped themes in embedding shadow roots ([5c4c5d7](https://github.com/ChristopherVR/ooxml/commit/5c4c5d7409e1a68915e679981e4a4e1943ba2e4a))
- **visio:** Honor themes inside application shadow roots ([a6d7738](https://github.com/ChristopherVR/ooxml/commit/a6d7738d25f03dbbaacbb381fd789775d9cb5c0d))
- **docx:** Inherit themes inside embedding shadow roots ([874caef](https://github.com/ChristopherVR/ooxml/commit/874caef766ca224a90dd2188e7c5a3beb9b82eb9))
- **xlsx:** Install scoped themes only when connecting ([638dfd0](https://github.com/ChristopherVR/ooxml/commit/638dfd055896784dac6664e4dddb03edd735746d))
- **pptx:** Inherit shared office themes across native viewers ([8ae64f7](https://github.com/ChristopherVR/ooxml/commit/8ae64f7f72ea70c323fef9b6c712697c6c54aeb6))
- **pptx:** Preserve product roles in the office theme bridge ([98e749a](https://github.com/ChristopherVR/ooxml/commit/98e749a9db7942edab244ce85eff1c5b6a1b9f77))

### Refactor

- **ui:** Share legacy product theme compatibility aliases ([49a16c9](https://github.com/ChristopherVR/ooxml/commit/49a16c99e6e7d743fd891433d62bdeb3c346b91b))

### Documentation

- **ui:** Define the shared editor theme contract ([207c055](https://github.com/ChristopherVR/ooxml/commit/207c055d34abd0c17c72fc69ef26aab69cba7f02))

### Testing

- **ui:** Guard legacy theme scope compatibility ([627ff95](https://github.com/ChristopherVR/ooxml/commit/627ff9525568e642dbf1310d052d479771f92be9))
- **ui:** Preserve shared theme styles between worker tests ([2aa8974](https://github.com/ChristopherVR/ooxml/commit/2aa8974457360e1d8d59fa3a4bcc4af0051764f5))
- **pptx:** Assert scoped office theme bridge capture ([a66991f](https://github.com/ChristopherVR/ooxml/commit/a66991f26ab79782d6ffa3c48ae7f4057bda9126))
- **ui:** Rerun setup files in shared workers ([#114](https://github.com/ChristopherVR/ooxml/issues/114)) ([c681768](https://github.com/ChristopherVR/ooxml/commit/c681768ef0b7fe0550e606409a7d3dd626d91d02))

### Dependencies

- **deps:** Update dev tools and fix svelte type checks ([#107](https://github.com/ChristopherVR/ooxml/issues/107)) ([550eac2](https://github.com/ChristopherVR/ooxml/commit/550eac24d469d14245a0ac71789e01c8d5fb50f9))

### Chores

- Integrate latest main before full verification ([3bec814](https://github.com/ChristopherVR/ooxml/commit/3bec814857ae675f31eba14fcec59a3db4e1cb77))

## [1.18.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.18.3) - 2026-10-20

### Dependencies

- **deps:** Update emf-converter to 4.21.0 ([#106](https://github.com/ChristopherVR/ooxml/issues/106)) ([641ae31](https://github.com/ChristopherVR/ooxml/commit/641ae316b7de07ed6f3c5158e8b154c991e1b719))

## [1.18.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.18.2) - 2026-10-20

### Bug Fixes

- **ui:** Render freeform arrows with shared endpoint sizing ([dff5678](https://github.com/ChristopherVR/ooxml/commit/dff567825cde43c4fa46c0bb90afef8eb9e12782))
- **pptx:** Preserve metrics during staged text builds ([388c1b2](https://github.com/ChristopherVR/ooxml/commit/388c1b25b621e5409dd55a483922d8f7928f023e))

### Testing

- **ui:** Align marker assertions with powerpoint dimensions ([e5c2f76](https://github.com/ChristopherVR/ooxml/commit/e5c2f76e99e012c29412171c23dd7ade80640b5b))

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))
- **deps:** Merge dependency update ([#95](https://github.com/ChristopherVR/ooxml/issues/95)) ([0926eac](https://github.com/ChristopherVR/ooxml/commit/0926eac231cb49749433fb7a9eac870cc7996e52))
- **deps:** Merge dependency update ([#96](https://github.com/ChristopherVR/ooxml/issues/96)) ([33b0ebb](https://github.com/ChristopherVR/ooxml/commit/33b0ebb6a16b4448d577c169c79187466d4148a3))
- **deps:** Merge dependency update ([#97](https://github.com/ChristopherVR/ooxml/issues/97)) ([8ef37ee](https://github.com/ChristopherVR/ooxml/commit/8ef37ee45829ad74631367cfd2d9c2c986536a60))

### Chores

- **deps-dev:** Bump tsup from 8.5.1 to 9.0.0 ([9835e80](https://github.com/ChristopherVR/ooxml/commit/9835e8038a9a271e37e286e524496513910553a6))
- **deps-dev:** Bump the development-minor-and-patch group with 14 updates ([1cdc60b](https://github.com/ChristopherVR/ooxml/commit/1cdc60b0f0a591304a2240ffcb550eccfd8d2487))
- **deps-dev:** Bump happy-dom from 20.14.6 to 21.0.3 ([b9bab0e](https://github.com/ChristopherVR/ooxml/commit/b9bab0e0ab3661d833feb74653788c81bde06c45))
- **pptx:** Integrate current main with dependency updates ([713239b](https://github.com/ChristopherVR/ooxml/commit/713239b3b9166d0010dcd058ea758c11ff2425c0))

## [1.18.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.18.1) - 2026-10-17

### Bug Fixes

- **pptx:** Bound image caches and dispose chart subscriptions ([0f2c004](https://github.com/ChristopherVR/ooxml/commit/0f2c0041b94c32293f8502bc299a4c18f586b3f3))

### Refactor

- **chart:** Share axis ranges and native gridline rendering ([8c26687](https://github.com/ChristopherVR/ooxml/commit/8c26687a7b3196955df602198528adc83fce259b))
- **chart:** Share pptx marker geometry with neutral rendering ([b246d63](https://github.com/ChristopherVR/ooxml/commit/b246d6315c32412298508d9b2a944b0c9bafc26b))

## [1.18.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.18.0) - 2026-10-17

### Features

- **ui:** Insert and edit charts as embedded objects in visio ([7fbea85](https://github.com/ChristopherVR/ooxml/commit/7fbea8563b1bfe5ca2b925c72fcdc5dbc0e23261))
- **ui:** Honour more visio options: language, spelling, tabs, toolbar ([95505e5](https://github.com/ChristopherVR/ooxml/commit/95505e5b7b9ab81529131a2f325b434332149d99))
- **ui:** Add the Mini Toolbar to visio ([d513a7e](https://github.com/ChristopherVR/ooxml/commit/d513a7e28329b41f2c9ba5daaec7638f78541100))
- **ui:** Add Live Preview to visio's style and theme galleries ([ddb7151](https://github.com/ChristopherVR/ooxml/commit/ddb7151622b8930111c030327e4767c2a10ddafe))
- **ui:** Add AutoRecover to visio's Save options ([b30dc43](https://github.com/ChristopherVR/ooxml/commit/b30dc43d8b181708640eb8c053755f539dd18242))
- **chart:** Draw data labels in the shared chart renderer ([32142ef](https://github.com/ChristopherVR/ooxml/commit/32142ef32c392548fdf5f79aa9d7f659356fac70))

### Bug Fixes

- **ui:** Keep gallery tiles intact when the visio language changes ([47455ef](https://github.com/ChristopherVR/ooxml/commit/47455ef95553dfa592c58ea65d24a0ff7aa756cd))
- **ui:** Stop galleries rebuilding their tiles and icons on every render ([35043f8](https://github.com/ChristopherVR/ooxml/commit/35043f833160d712d42b3b93f4dd6f8c4f8956e6))
- **ui:** Reject stale visio recovery offers ([5f4012c](https://github.com/ChristopherVR/ooxml/commit/5f4012c96e870b56444e41018c855a3200303651))

## [1.17.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.17.0) - 2026-10-17

### Features

- **ui:** Honour undo levels and ScreenTip style in visio options ([1ce62d3](https://github.com/ChristopherVR/ooxml/commit/1ce62d39fe7c7101f91ce93d4a34e612830946be))
- **ui:** Share the chart picker and add Insert > Chart to visio ([ab7f2ea](https://github.com/ChristopherVR/ooxml/commit/ab7f2ea0208f698b47def57f361b5a1f4eb457ba))

### Bug Fixes

- **ui:** Wire visio behaviour options and test them ([11bb768](https://github.com/ChristopherVR/ooxml/commit/11bb768187d87d55acff1f74f3809b48dba27924))
- **ui:** Keep the chart picker to declared design tokens ([406d92a](https://github.com/ChristopherVR/ooxml/commit/406d92a212de56fbd2751e80a81f5f4cb4ac1ae9))

### Testing

- **ui:** Drive the visio behaviour options through the dialog ([dc55882](https://github.com/ChristopherVR/ooxml/commit/dc55882dcf4dac95c42ef90abda4aa100bc49a1c))

## [1.16.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.16.0) - 2026-10-16

### Features

- **ui:** Choose a report filter with Values on rows ([b985bfb](https://github.com/ChristopherVR/ooxml/commit/b985bfb1b339fa254402b575d1b56fc3d9d94a13))
- **ui:** Caption a new PivotTable in the editor language ([7e7eb30](https://github.com/ChristopherVR/ooxml/commit/7e7eb30ece4ac17db7019089140a1decef3bce5b))
- **ui:** Translate the Teams elements in en, de, es, fr and zh-CN ([4bfb846](https://github.com/ChristopherVR/ooxml/commit/4bfb846c566bb24cd44fe789a64707df17adda51))
- **ui:** Add Effects and Background Styles galleries and hover preview ([fe19278](https://github.com/ChristopherVR/ooxml/commit/fe1927855182eae37051160f645ca6eff42c6b84))
- **ui:** Host the Design gallery preview and commit over deck state ([f3afeec](https://github.com/ChristopherVR/ooxml/commit/f3afeec39cbed7183e6b98606fcc6d6371dfc659))
- **ui:** Add Link and Date & Time to the phone Insert sheet ([f9962a5](https://github.com/ChristopherVR/ooxml/commit/f9962a5ddc0aebbaba8568f2203627c6989bc4c6))
- **ui:** Finish the chart Format pane's fills, 3-D and Text Options ([5bbd3db](https://github.com/ChristopherVR/ooxml/commit/5bbd3dbdab0c2177acfb3ba5afcbd682445359c4))
- **ui:** Draw chart picture, pattern, 3-D, text and line formats ([69b336b](https://github.com/ChristopherVR/ooxml/commit/69b336b7df631278a65ffb7ccefa67630430f847))

### Bug Fixes

- **ui:** Name the chart display units in words as office does ([206425d](https://github.com/ChristopherVR/ooxml/commit/206425d72c08a5ca01d8ef3986dc4aa6f2bdcf76))
- **ui:** Translate built-in SmartArt layout names in the Layouts gallery ([a48c035](https://github.com/ChristopherVR/ooxml/commit/a48c035a309cfadb46eb91e651839efaf559cf33))
- **ui:** Translate Compare change sentences and the waterfall legend ([f6f8e77](https://github.com/ChristopherVR/ooxml/commit/f6f8e77a0db1e2b418c5a326abae9b8a6a51124f))

### Testing

- **ui:** Refresh chart parity, manifest and locale checks after merge ([9bc34ea](https://github.com/ChristopherVR/ooxml/commit/9bc34eaafa0853af9e326c7ba5d0fea67aae80ef))
- **ui:** Keep the em-dash out of the teams locale test source ([bff7218](https://github.com/ChristopherVR/ooxml/commit/bff72183793478c22849c4201aa04770d8d028ba))

### Chores

- **ui:** Regenerate the element manifest for the gallery preview events ([916db87](https://github.com/ChristopherVR/ooxml/commit/916db8777efc2f2130a2c06d448b1d09022080f1))

## [1.15.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.15.0) - 2026-10-16

### Features

- **ui:** Add a sign-in screen to the teams app ([e6ccc71](https://github.com/ChristopherVR/ooxml/commit/e6ccc71008003355b99554441b1dc4a84d01e2fc))
- **ui:** Pick monthly and yearly repeats for a meeting ([bfc816a](https://github.com/ChristopherVR/ooxml/commit/bfc816a82c2c2670ae7f3dad7060bf37d86275b1))
- **ui:** Add the phone Insert sheet and a phone slide fit ([00d8320](https://github.com/ChristopherVR/ooxml/commit/00d83205847b26a985e72456f25e4ea78356ac2b))
- **ui:** Set Field Settings subtotals on PivotTable column fields ([71ae55a](https://github.com/ChristopherVR/ooxml/commit/71ae55a6314eae2ab2e32ebd4fa5e42886e8a70b))
- **ui:** Place PivotTable Values among any number of row fields ([9d3b3df](https://github.com/ChristopherVR/ooxml/commit/9d3b3df483d1276b913e42c9800c4a1668bfb65b))
- **ui:** Categorise the Shapes gallery and move Slide Templates ([e043c1a](https://github.com/ChristopherVR/ooxml/commit/e043c1a11c707277879617f9ae37d8901535c5a3))
- **ui:** Add the Themes and Variants galleries and Designer to Design ([75ec95b](https://github.com/ChristopherVR/ooxml/commit/75ec95b682e017657310bf5dbfc8b6a8a0de5585))
- **pptx:** Keep Slide Templates reachable from every New Slide menu ([be03af6](https://github.com/ChristopherVR/ooxml/commit/be03af63a0ed6c23a79800b500d3d425ea22773f))
- **ui:** Add the shared pptx chart Format task pane ([7cc8761](https://github.com/ChristopherVR/ooxml/commit/7cc8761e3c8f7ea3c55ab54a9760a19a2650960a))
- **ui:** Draw chart element transparency, effects and boxes ([d406438](https://github.com/ChristopherVR/ooxml/commit/d4064386a4e1d802de68fc0d829f887624ebc9cc))

### Bug Fixes

- **svelte:** Translate media and picture inspector labels ([d0bde63](https://github.com/ChristopherVR/ooxml/commit/d0bde63ea46502a316e94f4b795901db20644c17))
- **visio:** End presentation when the editor stops rendering ([e16d93e](https://github.com/ChristopherVR/ooxml/commit/e16d93ee5bb21c655ff30102b1d4d54bced9e12d))
- **visio:** Tidy the phone ribbon and narrow status bar ([32098eb](https://github.com/ChristopherVR/ooxml/commit/32098ebff2992dd305f60caa96d1154336f44226))
- **svelte:** Translate table, text effects and SmartArt inspector labels ([ca38b70](https://github.com/ChristopherVR/ooxml/commit/ca38b708db0399fec7db9f72046296c82587116d))
- **ui:** Translate shared preset catalogue labels ([46661a1](https://github.com/ChristopherVR/ooxml/commit/46661a125e90b91b00f3fa66f584f232bc21a871))
- **ui:** Keep chart Format pane clicks and target marks under capture ([8c47310](https://github.com/ChristopherVR/ooxml/commit/8c473101245fbc5635dc278826064cb8f52b138e))
- **ui:** End the phone chart Format sheet above the bottom bar ([16ea32a](https://github.com/ChristopherVR/ooxml/commit/16ea32a25a3f48537b3445ddab5e943b37a7a804))

### Testing

- **ui:** Group the dates of a PivotTable that has a report filter ([96efdb6](https://github.com/ChristopherVR/ooxml/commit/96efdb67138d46df04cf99c03d8dacdbdc53143e))

## [1.14.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.14.0) - 2026-10-15

### Features

- **ui:** Add a translator for the visio ribbon ([c79d403](https://github.com/ChristopherVR/ooxml/commit/c79d40343b8c503927f43bfb2d089971977baa87))
- **ui:** Fold visio's phone ribbon like word and excel ([871b77f](https://github.com/ChristopherVR/ooxml/commit/871b77fdb53a36bf37450c0e535c96a9cf2b06b2))
- **ui:** Track caret affinity on the print layout selection ([7bed9f4](https://github.com/ChristopherVR/ooxml/commit/7bed9f4f1cc276729888400cb19520c473dab869))
- **ui:** Tab and arrow through table cells on print layout pages ([43d7c81](https://github.com/ChristopherVR/ooxml/commit/43d7c816c8d6fa2003e0fde7753b247d8f25a7f8))
- **ui:** Edit headers and footers in place on print layout pages ([9e985ae](https://github.com/ChristopherVR/ooxml/commit/9e985ae1a29ec3a23d9555d2b5e847bf143c3173))
- **xlsx:** Draw combination, waterfall and histogram charts ([753dea7](https://github.com/ChristopherVR/ooxml/commit/753dea7e9f64cb3641643df907d5faa35612cf11))
- **ui:** Add excel page layout and page break preview views ([8cbb6e5](https://github.com/ChristopherVR/ooxml/commit/8cbb6e573aad3ee27c869e44fff2092578d05ec7))
- **ui:** Add what-if analysis to the excel data tab ([e51f768](https://github.com/ChristopherVR/ooxml/commit/e51f768738bd76d88d3122b2d32b5bd49a2b8cd5))
- **teams:** Add @mentions with a people picker in the composer ([2bb877a](https://github.com/ChristopherVR/ooxml/commit/2bb877a050dae027eb948e6d52b9b068934d927e))
- **teams:** Add an Activity feed derived from the shared chat data ([6f53e8a](https://github.com/ChristopherVR/ooxml/commit/6f53e8ac925ba74f7f4a4b1dec37c5b6849797ff))
- **teams:** Announce new activity with toasts, desktop notifications and sound ([fe29b00](https://github.com/ChristopherVR/ooxml/commit/fe29b006d4830489b19c097290b132a5581ddb98))
- **ui:** Drive powerpoint's options from the shared options dialog ([4ed1085](https://github.com/ChristopherVR/ooxml/commit/4ed1085fbdbb42519cdeb891ca5ce2624db4b1a1))
- **ui:** Open stencil files and collect My Shapes in the visio pane ([fe6312f](https://github.com/ChristopherVR/ooxml/commit/fe6312fa2da5e0dab162acacfc28b3b91fb80bc2))
- **teams:** Add private chats, meetings and a calendar to the client ([a6c2748](https://github.com/ChristopherVR/ooxml/commit/a6c2748426473192a2e16d37d3856f8accc27673))
- **ui:** Add Chat and Calendar apps to the Teams app ([2ddec01](https://github.com/ChristopherVR/ooxml/commit/2ddec01568a2ff8e79fbf9666f0f1eb0941d8cee))
- **ui:** Show a More sheet in the phone bottom bar ([a344495](https://github.com/ChristopherVR/ooxml/commit/a344495a4cb21c38eb70e7a79a96a5aa333f32f2))
- **ui:** Edit nested table cells on print layout pages ([8b2662b](https://github.com/ChristopherVR/ooxml/commit/8b2662b9dc87322311a02d7ec6e9b97164da6280))
- **ui:** Edit footnotes and endnotes on print layout pages ([294dd13](https://github.com/ChristopherVR/ooxml/commit/294dd131bf2701bec09f25ac5e9ca5a033bfc39a))
- **ui:** Select, move and resize floating pictures on pages ([95a0799](https://github.com/ChristopherVR/ooxml/commit/95a0799d9d3f262060a62dec0acf87c4c1141827))
- **docx:** Paginate plain documents incrementally after an edit ([7521712](https://github.com/ChristopherVR/ooxml/commit/7521712eba07913ff90a9205d982cd7a434a95b9))
- **ui:** Draw slicers and add Insert > Slicer with a Slicer tab ([452a351](https://github.com/ChristopherVR/ooxml/commit/452a3510399551fc80cb6f9b1f480217a71ccc0e))
- **ui:** Draw timelines and add Insert > Timeline with a Timeline tab ([c525b8c](https://github.com/ChristopherVR/ooxml/commit/c525b8c83558ba8c23ec0e2c772cae57bd4e26d8))
- **xlsx:** Group PivotTable fields by number ranges ([9a9014f](https://github.com/ChristopherVR/ooxml/commit/9a9014f0728979e09843d7781130e09b45cecfac))
- **docx:** Edit and insert content controls in the editor ([b1f3fce](https://github.com/ChristopherVR/ooxml/commit/b1f3fce94392d32919e1bd716c1cc5c0b7bf8911))
- **ui:** Insert and edit equations from LaTeX in the Word editor ([88b8964](https://github.com/ChristopherVR/ooxml/commit/88b8964bc1ff1509f37889f835dad75762f57268))
- **ui:** Draw and format shapes from Insert > Shapes ([b63e03c](https://github.com/ChristopherVR/ooxml/commit/b63e03c93b3dca7e9bb37946b9a841f7b779358c))
- **ui:** Let pptx hosts replace regions and set capabilities ([fefbde5](https://github.com/ChristopherVR/ooxml/commit/fefbde5c7385b1805d8d4824704fdefe28d591f4))
- **pptx:** Close binding props drift and guard prop parity ([167802e](https://github.com/ChristopherVR/ooxml/commit/167802e036f62e4c0dd23144a45ae1618eefd007))
- **ui:** Let the options dialog host a pptx region body ([76299e0](https://github.com/ChristopherVR/ooxml/commit/76299e01420f280f5881c91e96bbb763c6c4c31e))
- **ui:** Select, move, resize and edit floating shapes on pages ([b31198c](https://github.com/ChristopherVR/ooxml/commit/b31198c0fce3fa4910b5fde333daea3f0fae1ae2))
- **ui:** Recurrence end and this-event edits in the Teams calendar ([88f0c9f](https://github.com/ChristopherVR/ooxml/commit/88f0c9f88bd4129fd8d9bf9cf2219596316a8945))
- **ui:** Group dates, nest row fields and set subtotals per field ([614ac2c](https://github.com/ChristopherVR/ooxml/commit/614ac2c5206fb8b5a7fe538cccbb31ef958d248d))
- **ui:** Add further column fields in the PivotTable settings ([a8d540b](https://github.com/ChristopherVR/ooxml/commit/a8d540b4b5292349cca62ee21476d45352b1513e))
- **ui:** Lay out the PowerPoint Insert tab as Office does ([55aef8a](https://github.com/ChristopherVR/ooxml/commit/55aef8a30fd3ad7eabf0f2c4f1bd2fa3d8b0f49c))
- **ui:** Scale the Insert tab like Office and add the Table grid ([e6f5dcc](https://github.com/ChristopherVR/ooxml/commit/e6f5dcc7f050f184de46af1ed07bed16ff6dc446))

### Bug Fixes

- **ui:** Escape morph selector values against markup injection ([08f3a94](https://github.com/ChristopherVR/ooxml/commit/08f3a94bc1c2916dd80bb5588c109a50005189fe))
- **ui:** Escape the flood colour in the image fill overlay filter ([f2a0bbb](https://github.com/ChristopherVR/ooxml/commit/f2a0bbbe6cb851e2e182a9366fb3611b5045545d))
- **ui:** Sandbox the xlsx print preview and print frames ([f4b87c9](https://github.com/ChristopherVR/ooxml/commit/f4b87c956d49a5cfbffcfc3783175a2ae8f49077))
- **ui:** Strip active content from docx chart svg before printing ([53849cd](https://github.com/ChristopherVR/ooxml/commit/53849cd6ed1225518625dc0292b06013ab226b92))
- **ui:** Refuse markup MIME types when opening OLE payloads in a tab ([4e997c8](https://github.com/ChristopherVR/ooxml/commit/4e997c8975c2f189f8cd1cfccded65920c588de7))
- **ui:** Align page layout rulers and fit the page to narrow screens ([713f952](https://github.com/ChristopherVR/ooxml/commit/713f95291e7dc138f7d41a3fb935606d6546c677))
- **teams:** Place notification toasts clear of the composer and show mentions without @ ([1e46d6d](https://github.com/ChristopherVR/ooxml/commit/1e46d6dc907b72cbe63359ae94cd4648b5e09697))
- **ui:** Ignore edits to locked powerpoint options in the shared dialog ([4e9a0df](https://github.com/ChristopherVR/ooxml/commit/4e9a0df07a2867636f4b62882281df0f55643a31))
- **ui:** Theme the shared options dialog's native controls ([85ad513](https://github.com/ChristopherVR/ooxml/commit/85ad513e509c2df44185ae4c280221790fe4d6b9))
- **ui:** Keep the default colour scheme when tokens cannot resolve ([d34c6ab](https://github.com/ChristopherVR/ooxml/commit/d34c6abd43879aeeadb79169262b8ef7d68c9f68))
- **ui:** Tolerate excel sheets without a view when reading the mode ([6ffb621](https://github.com/ChristopherVR/ooxml/commit/6ffb62192685f0156006077c8398de7c1aa3d452))
- **ui:** Polish the Chat and Calendar apps after a browser pass ([02e5543](https://github.com/ChristopherVR/ooxml/commit/02e5543393713280ab8c941094ef4a282c39ddfc))
- **ui:** Keep the page of an open story editor mounted when scrolling ([7e441d0](https://github.com/ChristopherVR/ooxml/commit/7e441d0567b9e58998454d2478650d84ebf3fadf))
- **xlsx:** Soften slicer and timeline style tints ([5640e20](https://github.com/ChristopherVR/ooxml/commit/5640e20a4345f47e2b6980858c80e37ff13d7f68))
- **ui:** Keep powerpoint ribbon tabs whole in a narrow editor ([604294a](https://github.com/ChristopherVR/ooxml/commit/604294a1b9b9e18abf0bd1c67a39a7aeabe1c473))
- **ui:** Correct visio ribbon translations ([c21a990](https://github.com/ChristopherVR/ooxml/commit/c21a990f83a585593528f65778369f8d02895832))
- **svelte:** Translate the chart section and lead with its title ([1a102b6](https://github.com/ChristopherVR/ooxml/commit/1a102b60aeb7def69f05fd2d5d0b18286467cd46))
- **ui:** Keep unchanged visio guides across renders ([3a4fa26](https://github.com/ChristopherVR/ooxml/commit/3a4fa26acdefc27aaf48f6e6f84a41b27ee47e73))

### Refactor

- **chart:** Move waterfall steps and range into the chart area ([b6fa22a](https://github.com/ChristopherVR/ooxml/commit/b6fa22a987d40b63f3369868083435eba35afa83))
- **chart:** Move number formats, palettes, axis labels, auto title to core ([26bde03](https://github.com/ChristopherVR/ooxml/commit/26bde036faed3663e9a037b8e755e5c517202263))
- **chart:** Move bubble, stock, treemap and sunburst geometry to core ([a465fa6](https://github.com/ChristopherVR/ooxml/commit/a465fa63dfd1f15c67a22a97fa5ed9c97d2965e8))
- **ui:** Draw pptx trendlines and error bars through core fits ([ea36c2b](https://github.com/ChristopherVR/ooxml/commit/ea36c2b29c43d13596829418d96a9c39f095d12e))

### Documentation

- **xlsx:** Record slicers, timelines and number groups ([a397273](https://github.com/ChristopherVR/ooxml/commit/a3972734b2b384365cf13a73d5f475fec47187ee))

### Testing

- **xlsx:** Cover page views in the browser and document them ([cc2eec8](https://github.com/ChristopherVR/ooxml/commit/cc2eec8d9b641d8b1d9673f04e1da61d40580a95))
- **pptx:** Guard chart view models with a golden parity snapshot ([9100a2b](https://github.com/ChristopherVR/ooxml/commit/9100a2b361c764761c51a7a0c113bdd2b89d1453))
- **ui:** Count floating shapes when locating a picture on a page ([ad97f2f](https://github.com/ChristopherVR/ooxml/commit/ad97f2f1c94ef0adc27c93ccc58e622e23a5d782))
- **ui:** Remove the clipboard the word context menu test defines ([260a953](https://github.com/ChristopherVR/ooxml/commit/260a953eabdd657fb744b8c4881d03f51da27a45))
- **ui:** Give the ssr import tests room under load ([ee53704](https://github.com/ChristopherVR/ooxml/commit/ee5370423a0075a8a53c48557baef24d120c7db2))
- **ui:** Unstub the globals the webfont and image tests replace ([d495e75](https://github.com/ChristopherVR/ooxml/commit/d495e75c188d605fd0a483a6dc3632e4b3c65d3c))

### Chores

- **ui:** Drop an unused chart import and clarify the note lookup ([76ed6f1](https://github.com/ChristopherVR/ooxml/commit/76ed6f1bbd263129a0d3950119a8f062ba020d9f))
- Format files touched by the round three merges ([1da72f9](https://github.com/ChristopherVR/ooxml/commit/1da72f9bc3f5bc20adff045c183048b31bc58b37))

## [1.13.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.13.0) - 2026-10-15

### Features

- **ui:** Fit word pages to narrow screens and trim status bars ([920e449](https://github.com/ChristopherVR/ooxml/commit/920e449c6a9c008ef9ec779bfbfd8d49056c3cbd))
- **ui:** Add touch pinch zoom and long-press to the visio editor ([9e6acfd](https://github.com/ChristopherVR/ooxml/commit/9e6acfda8c72415bb4fb6b168243ddefb06a64ff))
- **ui:** Add touch and narrow modes to the shared ribbon ([fa77402](https://github.com/ChristopherVR/ooxml/commit/fa77402f743242ef8c389f0ae0eccc6cd345b1e1))
- **ui:** Add teams mobile navigation with bottom bar and list view ([b68c472](https://github.com/ChristopherVR/ooxml/commit/b68c472459da73dc2d40b3325122f84d619c3194))
- **teams:** Pin and unpin channel messages with a pinned bar ([00a4c60](https://github.com/ChristopherVR/ooxml/commit/00a4c603e4b95be5242ec30257a64e4355bc76c4))
- **teams:** Add do not disturb, status messages and idle away ([e0adea1](https://github.com/ChristopherVR/ooxml/commit/e0adea1e772fd89eaec14c87b4874823d62b0912))
- **ui:** Add pptx table layout tab and review comment navigation ([325ab16](https://github.com/ChristopherVR/ooxml/commit/325ab1613592c633a6d1291ecee7d05478e40485))
- **ui:** Lay out the shared options dialog like powerpoint's ([a999f52](https://github.com/ChristopherVR/ooxml/commit/a999f526af7aac2c4758e421e939ed0738b694f0))
- **ui:** Show word's options in the shared options dialog ([9edec67](https://github.com/ChristopherVR/ooxml/commit/9edec67bb2f6698bccad8ed5fc79115bb9c0c583))
- **ui:** Show excel's options in the shared options dialog ([c5db5ed](https://github.com/ChristopherVR/ooxml/commit/c5db5ed7cf41314bd2e2e5dfef5c9a8d12e1c2fe))

### Bug Fixes

- **ui:** Replace the stale visio position menu reason ([9a0eec3](https://github.com/ChristopherVR/ooxml/commit/9a0eec37c46167ba2fedd090a7f90d989ecedcf6))
- **ui:** Keep the word page counter readable on narrow status bars ([0c4ba89](https://github.com/ChristopherVR/ooxml/commit/0c4ba8985fe4e5f07480baa37e87e58358cce9e2))
- **ui:** Translate the ribbon scroll chevron label ([b894992](https://github.com/ChristopherVR/ooxml/commit/b894992014f0aa21f6922f8a55da6db2b12b122e))
- **ui:** Regenerate the element manifest and tokenise the rail insets ([e39033e](https://github.com/ChristopherVR/ooxml/commit/e39033e5501188393e5dcafcc332338c19491a50))
- **ui:** Place the print caret at the nearest line on any page click ([dc6d0e1](https://github.com/ChristopherVR/ooxml/commit/dc6d0e128b305de8e5b426f6aa40811d6f75b28c))
- **ui:** Put the caret at the nearest text when the bare canvas is pressed ([5f39bad](https://github.com/ChristopherVR/ooxml/commit/5f39bad513bf0ed65581ccb488e10d66241fb674))
- **ui:** Zoom Print Layout once and stop the post-drag scroll jump ([0dba74e](https://github.com/ChristopherVR/ooxml/commit/0dba74e2e359be7f2859ab999585c260aa8b942b))
- **ui:** Match the options dialog footer, size and focus to powerpoint's ([fda0178](https://github.com/ChristopherVR/ooxml/commit/fda01784fd02d6df9fb7ff9b46eba5ed98763d4f))
- **ui:** Lay out print view at once when a document is swapped ([b8df7e4](https://github.com/ChristopherVR/ooxml/commit/b8df7e4d846529bc81b9f58055111dc519ae0b09))

### Refactor

- **ui:** Split visio pointer gestures and test touch hit tolerance ([260da9f](https://github.com/ChristopherVR/ooxml/commit/260da9fca8d387f7a108b244086621e4b9b56701))
- **ui:** Build visio's options categories from the shared ones ([559300d](https://github.com/ChristopherVR/ooxml/commit/559300dfcd5bb0259c0a185be3cf9174a634db58))

### Testing

- **ui:** Format the options dialog test ([1e49d9e](https://github.com/ChristopherVR/ooxml/commit/1e49d9ec330de11fb73e29527bbb63fd885d7873))

## [1.12.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.12.3) - 2026-10-14

### Bug Fixes

- **ui:** Translate swatch, style and scheme names in pptx menus ([#83](https://github.com/ChristopherVR/ooxml/issues/83)) ([bd2908c](https://github.com/ChristopherVR/ooxml/commit/bd2908c03147486699378b99eccbba6cbd5b8efd))
- **ui:** Draw the table background and lines of Themed Styles ([#76](https://github.com/ChristopherVR/ooxml/issues/76)) ([8ea2382](https://github.com/ChristopherVR/ooxml/commit/8ea2382e8654a12e5f8f48d1658c8e872b7c9cc0))

## [1.12.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.12.2) - 2026-10-14

### Performance

- **docx:** Cache review data during caret updates ([e03f8c8](https://github.com/ChristopherVR/ooxml/commit/e03f8c82fdb7418e84bc0c041139e969c1633e17))

## [1.12.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.12.1) - 2026-10-14

### Bug Fixes

- **ui:** Restore ribbon layout and stabilize browser checks ([cba6254](https://github.com/ChristopherVR/ooxml/commit/cba625414fec57c00fa929d19ff8f4337af9a4f5))
- **pptx:** Commit inline edits and suppress bookmark shortcuts ([fea149f](https://github.com/ChristopherVR/ooxml/commit/fea149fff0444996b68d9c869d5209acdc9581ea))
- **pptx:** Close outline view and format background on escape ([8b837d7](https://github.com/ChristopherVR/ooxml/commit/8b837d71ee66588e7912a958c891c31e4e19868a))

## [1.12.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.12.0) - 2026-10-14

### Features

- **ui:** Align word excel and visio toolbar surfaces ([9decf56](https://github.com/ChristopherVR/ooxml/commit/9decf563e3c58be4fd7c767a4a459fad3a45029f))

### Bug Fixes

- **ui:** Restore native touch dragging across office editors ([6dc56a1](https://github.com/ChristopherVR/ooxml/commit/6dc56a1ba920b230d2a2d23572378305e74fe48c))
- **ui:** Correct mobile interaction test fixtures ([14dd76f](https://github.com/ChristopherVR/ooxml/commit/14dd76f4e5c2f988b0d145962523a564702201f1))

### Performance

- **docx:** Reduce typing and print layout work ([5b06c49](https://github.com/ChristopherVR/ooxml/commit/5b06c49ece0ecb00bd24c40db9df6ac7777e3525))

### Testing

- **docx:** Cover native table keeping controls and history ([ccefed0](https://github.com/ChristopherVR/ooxml/commit/ccefed0e453d1b880bd70d0f6f57d75fca2300c9))

## [1.11.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.11.0) - 2026-10-13

### Features

- **ui:** Expose existing visio field editing in the dialog ([0a66e1c](https://github.com/ChristopherVR/ooxml/commit/0a66e1c00a30f6ca5609ae51c3caeef36bb6a942))
- **xlsx:** Configure cached report filter selections ([993a5a5](https://github.com/ChristopherVR/ooxml/commit/993a5a5922656ccf90f8eece099a804200a7bc43))
- **ui:** Customize inspector startup and deck card visibility ([ada4979](https://github.com/ChristopherVR/ooxml/commit/ada49799906e5ed99159b1f326ae0cb1c737017e))
- **ui:** Configure nested pivot report filters ([6b3ec8a](https://github.com/ChristopherVR/ooxml/commit/6b3ec8af611429ce8bf75d7d0ef91ed2eae574ec))

### Bug Fixes

- **pptx:** Honor bounded native visibility timing ([a9a70ee](https://github.com/ChristopherVR/ooxml/commit/a9a70ee8c87abfb32ab3d6999c633073bf5e3ddd))
- **visio:** Restrict existing field choices to local sources ([4b718d5](https://github.com/ChristopherVR/ooxml/commit/4b718d5d178485377044635f9f110d583aabdfe0))
- **ui:** Defer title bar search hooks until rendering ([8a78d7d](https://github.com/ChristopherVR/ooxml/commit/8a78d7d082b4f54f2ff30558abc0463160f5b47a))
- **pptx:** Separate inherited paragraph spacing from direct values ([0c23aac](https://github.com/ChristopherVR/ooxml/commit/0c23aac22831cff4dd4a7dcce79d48ffdaab27b4))
- **ui:** Describe measured pivot report filter limits ([e3adfc8](https://github.com/ChristopherVR/ooxml/commit/e3adfc800a994de92d717bcc8cfa3d5aa62be0d4))
- **ui:** Open Visio page setup from context menu ([9ba68d4](https://github.com/ChristopherVR/ooxml/commit/9ba68d4fb3c2e449bc1ac50ac0e8fa215f652327))
- **pptx:** Apply explicit line spacing over inherited paragraphs ([aad8c3d](https://github.com/ChristopherVR/ooxml/commit/aad8c3deba0eb8d4034b208e732136a5fc24cc92))
- **ui:** Declare and document select reapply metadata ([820df4b](https://github.com/ChristopherVR/ooxml/commit/820df4b5130a893b11e23cce63453437e53c72df))

### Refactor

- **pptx:** Move native effect timing decisions into core ([61e030a](https://github.com/ChristopherVR/ooxml/commit/61e030a0ae01032d36da6211577bf340187052c8))
- **pptx:** Share font and whole-text edits in core ([94be138](https://github.com/ChristopherVR/ooxml/commit/94be138efc09b3eb881436eb07be54af0460fb45))
- **pptx:** Move paragraph and run style scope into core ([3ba3c32](https://github.com/ChristopherVR/ooxml/commit/3ba3c321af80efed2da4c6d58a377070c59b7c07))
- **pptx:** Move paragraph geometry cascade into core ([2dfc204](https://github.com/ChristopherVR/ooxml/commit/2dfc204f3ae6a84fed476073b3aefae261ab8d6b))

### Testing

- **docx:** Verify two-row nested print workflows ([d4514b9](https://github.com/ChristopherVR/ooxml/commit/d4514b9eca4d70fbd032747e5e477fb54782a575))
- **pptx:** Preserve inherited spacing through core body edits ([29c3b46](https://github.com/ChristopherVR/ooxml/commit/29c3b469141f3d338cfc2f8af44aea80f1963807))
- **docx:** Verify native three-row print workflows ([b2414f7](https://github.com/ChristopherVR/ooxml/commit/b2414f77b1bac171f73948f0aec8b6a173888d1d))
- **docx:** Verify structural nested row print workflows ([b269d5b](https://github.com/ChristopherVR/ooxml/commit/b269d5bd94bff77850327e4ca9bc3514f2061430))
- **ui:** Assert measured pivot help bounds ([51e7641](https://github.com/ChristopherVR/ooxml/commit/51e76418c8428a3101d93a0d3f24df193bbf05ec))

### Dependencies

- **deps:** Update production dependencies ([bbd7b53](https://github.com/ChristopherVR/ooxml/commit/bbd7b530c7a15ae48c0c0a2ff4ea7dcdfe2aa9d3))

## [1.10.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.10.0) - 2026-10-13

### Features

- **ui:** Configure multiple pivot value fields ([eede3cb](https://github.com/ChristopherVR/ooxml/commit/eede3cb1b058435cac114e716db834ffe5757685))
- **ui:** Enable multi-measure pivot report layouts ([bf559ec](https://github.com/ChristopherVR/ooxml/commit/bf559ec8eb746c7b3a1d1869d04e0a4a51467c94))
- **ui:** Expose pivot values axis settings ([9a6164f](https://github.com/ChristopherVR/ooxml/commit/9a6164f7e0fa3ed23b337a601eba0524c4921323))
- **ui:** Add a Size group with Crop to the Picture Format tab ([58f5906](https://github.com/ChristopherVR/ooxml/commit/58f5906b157cf2503b9348f671da76c591117cc6))
- **ui:** Enable compact and outline pivot row values ([87535e3](https://github.com/ChristopherVR/ooxml/commit/87535e36cbf34ff04e0811d2db8ab70f15268580))
- **pptx:** Support native whole-shape fixed opacity ([7b0372e](https://github.com/ChristopherVR/ooxml/commit/7b0372ed918da963bf662a302334e68ccae972e2))
- **xlsx:** Author nested values rows in tabular reports ([fc94245](https://github.com/ChristopherVR/ooxml/commit/fc94245e3f5f2caf6529a2daabf4f0e25229c47e))
- **xlsx:** Configure nested row values in all report layouts ([f85cf7d](https://github.com/ChristopherVR/ooxml/commit/f85cf7d8c2580d5b4bc1b0fb8a0451efe577dc9a))
- **xlsx:** Configure between-field values placement ([07357bd](https://github.com/ChristopherVR/ooxml/commit/07357bde39824b7f433be2d6b7e7529b1ed041b0))
- **xlsx:** Configure values before row fields ([4f33bd0](https://github.com/ChristopherVR/ooxml/commit/4f33bd0d2e27dd113ef21aeead69b8940a70b8bb))

### Bug Fixes

- **pptx:** Paint animated picture frames above their images ([55ac6e8](https://github.com/ChristopherVR/ooxml/commit/55ac6e8d9ae5b4c74d0f1bd18b783864c9d446c1))
- **docx:** Paint nested table frames in print layout ([0415699](https://github.com/ChristopherVR/ooxml/commit/04156990666eb0a9f7b449c78266000578ea3797))
- **docx:** Align nested print text with cell frames ([66a9030](https://github.com/ChristopherVR/ooxml/commit/66a9030394475733f58f06541206274e27652c15))
- **docx:** Reserve measured nested table border insets ([d59a6ea](https://github.com/ChristopherVR/ooxml/commit/d59a6ea69bad31c395f58caed8160ebac082209f))
- **docx:** Measure parallel table continuation bottoms ([854baa9](https://github.com/ChristopherVR/ooxml/commit/854baa92664b927016425e8b1edaf7148d7d387b))
- **pptx:** Animate shape backgrounds independently of text ([1c2568b](https://github.com/ChristopherVR/ooxml/commit/1c2568b9fc5712ae34b9e0a11f530d4b298eb8a7))
- **pptx:** Preserve background paint colors and whole opacity ([afd6e3b](https://github.com/ChristopherVR/ooxml/commit/afd6e3b5e372a5d05c723e1cdb55c6c001f241dc))
- **ui:** Keep spaces and line breaks in a wrap="none" text body ([12297d1](https://github.com/ChristopherVR/ooxml/commit/12297d13787c6230e406ab52d53c07365bcae185))
- **docx:** Align asymmetric nested borders and paint layers ([7540cb1](https://github.com/ChristopherVR/ooxml/commit/7540cb1e6224298dec801134537512383aee007b))
- **docx:** Reserve ordinary table border occupancy ([b7c904e](https://github.com/ChristopherVR/ooxml/commit/b7c904edc474c2704075a97ad029890c42b46d7c))
- **pptx:** Preserve native after-state and honor explicit clears ([aed6a6d](https://github.com/ChristopherVR/ooxml/commit/aed6a6d03410a3487589edfa6f57bde243885a45))
- **pptx:** Number slide-number fields from the deck's firstSlideNum ([ec9b18d](https://github.com/ChristopherVR/ooxml/commit/ec9b18da409c971d08491b612db228fa32608c8b))
- **pptx:** Drop cleared cell fills and borders when the style changes ([99dbeba](https://github.com/ChristopherVR/ooxml/commit/99dbebabe7fd885998c52360610781e10ad3bca4))
- **ui:** Match table style previews to the table's style options ([2fcf53c](https://github.com/ChristopherVR/ooxml/commit/2fcf53c1eb19488b0a7873761a43b9df51532e01))
- **ui:** Apply a picked font family to the whole text box ([5c20e49](https://github.com/ChristopherVR/ooxml/commit/5c20e4945a8b35bf68169715f3f4bd926ecfac86))

## [1.9.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.9.0) - 2026-10-13

### Features

- **docx:** Improve footnote flow and print editing ([2fff641](https://github.com/ChristopherVR/ooxml/commit/2fff64125254edb739e874101007b189a1609a19))
- **xlsx:** Add sparkline pivot and print workflows ([bb34ffb](https://github.com/ChristopherVR/ooxml/commit/bb34ffb060b12049a780b68fc3c0ff3db0b4682a))
- **visio:** Match native routing and shape data behavior ([be1519c](https://github.com/ChristopherVR/ooxml/commit/be1519c203226dd531a3b147b9f2a532cd1945ba))
- **pptx:** Add rich search and cross-deck slide clipboard ([736dbf9](https://github.com/ChristopherVR/ooxml/commit/736dbf94f7fac5226accea2e942db2e5f9e08063))
- **docx:** Support native footnote columns across compatibility modes ([8033fec](https://github.com/ChristopherVR/ooxml/commit/8033fece4e1f8030216fca267939aea4587f2fc5))
- **xlsx:** Add hierarchical pivots and sparkline date axes ([81001cd](https://github.com/ChristopherVR/ooxml/commit/81001cd74e185a4a507e0b1524f1b1998333273a))
- **pptx:** Edit rich chart titles across viewer bindings ([e27970e](https://github.com/ChristopherVR/ooxml/commit/e27970e7791e8b1d2287619aa338ac907e25758d))
- **xlsx:** Author compact and outline pivot layouts ([611f377](https://github.com/ChristopherVR/ooxml/commit/611f37708c99b86e2a220f3c90458c54154a3c8b))
- **xlsx:** Render rich print header formatting ([ca16fe3](https://github.com/ChristopherVR/ooxml/commit/ca16fe39e5bc39cecf82ce57611f3d2062147d5b))
- **pptx:** Replace chart labels with coherent workbook updates ([df616f7](https://github.com/ChristopherVR/ooxml/commit/df616f7f14674af9e41fd579b6860280ef45b9ac))

### Bug Fixes

- **ui:** Round the table cell panel numbers for display ([#65](https://github.com/ChristopherVR/ooxml/issues/65)) ([964a96f](https://github.com/ChristopherVR/ooxml/commit/964a96f3c9c04e5bb186b1459d9089d8e6cf06ee))
- **docx:** Paginate footnotes referenced inside body tables ([e137736](https://github.com/ChristopherVR/ooxml/commit/e137736a5c11cc791f4a02da87e3f517205de1e4))
- **docx:** Retain cell padding when splitting table footnotes ([0e36444](https://github.com/ChristopherVR/ooxml/commit/0e36444a381f527d8cd03729e7bedd079ef3d29e))
- **docx:** Paginate nested table notes and guard source structure ([f67b8c3](https://github.com/ChristopherVR/ooxml/commit/f67b8c3491464230a3f8fd9dddf7654615d37af5))
- **ui:** Paint picture frames and shadows across bindings ([#69](https://github.com/ChristopherVR/ooxml/issues/69)) ([c012f08](https://github.com/ChristopherVR/ooxml/commit/c012f08aa95ca069f2164f9f50e15208a793ba8e))
- **pptx:** Handle opaque share origins and clean up dialog tests ([2c73b55](https://github.com/ChristopherVR/ooxml/commit/2c73b55ed37381c6b7f51cefe72bd2439d5e03af))

### Testing

- **docx:** Isolate alignment history from jsdom scrolling ([dd8c521](https://github.com/ChristopherVR/ooxml/commit/dd8c521407e97c21ee3273bd1002bef70c3652ee))

### Chores

- **deps-dev:** Bump the development-minor-and-patch group across 1 directory with 3 updates ([#60](https://github.com/ChristopherVR/ooxml/issues/60)) ([630adbb](https://github.com/ChristopherVR/ooxml/commit/630adbb3aa8d0a22007233ae64b999edf559115c))

## [1.8.5](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.8.5) - 2026-10-13

### Bug Fixes

- **ui:** Apply line spacing and indent over a paragraph's own a:pPr ([#62](https://github.com/ChristopherVR/ooxml/issues/62)) ([d8a24e6](https://github.com/ChristopherVR/ooxml/commit/d8a24e6740d773e4e4c49c8089f7ca4db20b3b3b))
- **ui:** Give a shape with no line a default width on outline colour ([#67](https://github.com/ChristopherVR/ooxml/issues/67)) ([2db8c24](https://github.com/ChristopherVR/ooxml/commit/2db8c2451264edc18a0c79d8a62fbd268fee6111))

## [1.8.4](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.8.4) - 2026-10-11

### Testing

- **ui:** Clear the text selection between test files ([abb829f](https://github.com/ChristopherVR/ooxml/commit/abb829f2e3a4e6ff0a99a9023012d1d3ad0d8b51))
- **ui:** Measure the autofit clone against the body it started with ([0336f59](https://github.com/ChristopherVR/ooxml/commit/0336f5963f4422e405cd7a4ad4206ad395643949))
- **ui:** Compare body children to the starting count in two more tests ([cd39154](https://github.com/ChristopherVR/ooxml/commit/cd3915465a0192b2a6fa1dafde17a77d63c31316))

## [1.8.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.8.3) - 2026-10-11

### Bug Fixes

- **ui:** Refresh the Layouts gallery when the layout library loads ([817daaf](https://github.com/ChristopherVR/ooxml/commit/817daaf54f1f58d8d7b6a44d1c65ddeecb17591a))
- **ui:** Do not offer a page the autosave snapshot it wrote itself ([e6b2e1a](https://github.com/ChristopherVR/ooxml/commit/e6b2e1a42c58b349a33e8cc17e77918ae96373c4))

### Testing

- **ui:** Keep each DOM environment and test state in its own worker ([6d52bec](https://github.com/ChristopherVR/ooxml/commit/6d52bec41343129aa621e11e173fe3c706e82b6e))

### Styling

- Fix the oxfmt and oxlint warnings outside the viewers ([cd66f32](https://github.com/ChristopherVR/ooxml/commit/cd66f324a45aad1d9bd680d1518f0dbe8baf8fc3))

## [1.8.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.8.2) - 2026-10-11

### Bug Fixes

- **ui:** Let bridged themes re-derive the tokens that read them ([531dae6](https://github.com/ChristopherVR/ooxml/commit/531dae6f2fe1cf31894d54b46df31a1de934d954))
- **ui:** Keep visio drag previews until the source edit settles ([a004a29](https://github.com/ChristopherVR/ooxml/commit/a004a292913c4ec63fe6aa6e21196c36eeef9dff))

## [1.8.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.8.1) - 2026-10-10

### Testing

- **ui:** Give jsdom viewer tests time to settle under parallel load ([239f9ab](https://github.com/ChristopherVR/ooxml/commit/239f9abf933d35fe13220acc248c6bffd2464715))
- **ui:** Share workers across files and run one vitest process ([0c7a8e5](https://github.com/ChristopherVR/ooxml/commit/0c7a8e5aa0998d79c5fe6224264a5035a54dc773))

## [1.8.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.8.0) - 2026-10-10

### Features

- **ui:** Let a host add ribbon tabs to the Visio, Excel and Word UI ([101c552](https://github.com/ChristopherVR/ooxml/commit/101c552196d3b9e2befa280b8ed795e4cf8a04b8))
- **visio:** Match Visio's ribbon tabs and title bar layout ([f57ee8e](https://github.com/ChristopherVR/ooxml/commit/f57ee8e649314514faf7f7656cdb7ca719967eea))
- **visio:** Show Quick Styles, Themes and Variants inline in the ribbon ([30baa65](https://github.com/ChristopherVR/ooxml/commit/30baa655815a974c41c5947e51d9809f3939ef40))
- **visio:** Collapse ribbon groups into buttons in a narrow window ([a9fd1a1](https://github.com/ChristopherVR/ooxml/commit/a9fd1a1b343068cca5436a0c05539ccd9b3419dc))
- **ui:** Add office-ui-color-grid, Office's colour picker body ([6d8a8aa](https://github.com/ChristopherVR/ooxml/commit/6d8a8aaa9d9aedd74428b8d37e56286896d9ca05))
- **visio:** Add Office's colour picker and a Format Shape task pane ([021f43b](https://github.com/ChristopherVR/ooxml/commit/021f43bb4267b4022cd7b771d03aadf2279a6d55))
- **visio:** Add AutoConnect arrows with a Quick Shapes mini toolbar ([55c5f4b](https://github.com/ChristopherVR/ooxml/commit/55c5f4bb7652d7f3ea70515357ab4815b8461b72))
- **visio:** Show the drawing's own stencils in the Shapes window ([913edfa](https://github.com/ChristopherVR/ooxml/commit/913edfa639320bb7c5eb0eb6dea91c794dc1e999))
- **visio:** Offer resize, move, rotate and formatting on stencil shapes ([6795c0e](https://github.com/ChristopherVR/ooxml/commit/6795c0ed6106f781178a49aecefd687993998479))
- **ribbon:** Add a core area for the ribbon add-in tab descriptor ([7b2fa48](https://github.com/ChristopherVR/ooxml/commit/7b2fa48d2562d7483e03acf573e74404342b8f2a))
- **visio:** Make ribbon add-in tabs a prop and event of every binding ([6e6a157](https://github.com/ChristopherVR/ooxml/commit/6e6a157bbced977145d428e2ca01168e0a413deb))
- **docx:** Make ribbon add-in tabs a prop and event of every binding ([819c1a3](https://github.com/ChristopherVR/ooxml/commit/819c1a38c2beb9605676d29994b71a7a6f7e4a15))
- **xlsx:** Make ribbon add-in tabs a prop and event of every binding ([691f553](https://github.com/ChristopherVR/ooxml/commit/691f5531b9f5b4a5a5ac0645b43c8c0fc55be5d8))
- **ui:** Add the shared pieces of PowerPoint ribbon add-in tabs ([70a3376](https://github.com/ChristopherVR/ooxml/commit/70a3376eaf1cc1bdbf1208baa626e2a850318261))
- **visio:** Draw line jumps and toggle them from Design > Connectors ([842ddbb](https://github.com/ChristopherVR/ooxml/commit/842ddbb2f64602a3c201678d73bb01aa45cfd9d7))
- **visio:** Add the Ruler & Grid dialog, saved to the page ([f6646a8](https://github.com/ChristopherVR/ooxml/commit/f6646a84815af9148b7e18351eb4b2488b47d6d4))
- **visio:** Add the Snap & Glue dialog, saved to the drawing ([f96c927](https://github.com/ChristopherVR/ooxml/commit/f96c927906a6e4d6872b4667439726460d91dd61))
- **visio:** Insert a horizontal legend from Data > Insert Legend ([92ac505](https://github.com/ChristopherVR/ooxml/commit/92ac5055f1325d1bff1472d822d31d027d80455c))
- **visio:** Edit stencil groups and their parts in the viewer ([977bc39](https://github.com/ChristopherVR/ooxml/commit/977bc397b2255bfc8e10462dadfec63544a3816c))
- **visio:** Measure text in the edit worker, grow the editing frame ([e54f715](https://github.com/ChristopherVR/ooxml/commit/e54f715afe410b209e2d2c01bfc0d68089c4a177))
- **visio:** Offer glue and end handles on stencil shapes and connectors ([7d815b9](https://github.com/ChristopherVR/ooxml/commit/7d815b9f29508f5427a70abcab7728d9a57a785d))
- **visio:** Drop the Dynamic connector from the Document Stencil ([93eb00a](https://github.com/ChristopherVR/ooxml/commit/93eb00ae70b5ffd830c7644a1910fac39faf7b68))
- **visio:** Offer whole-shape commands on stencil shapes ([4734035](https://github.com/ChristopherVR/ooxml/commit/4734035524b0cac0f084dab62db406d2d01e3f07))
- **visio:** Pick theme colours that follow the page theme ([d29c9c7](https://github.com/ChristopherVR/ooxml/commit/d29c9c7b8ae343821db3461ca55af2283d6a19ef))
- **ui:** Add office-ui-color-custom, the Colors dialog's Custom tab ([ae061fd](https://github.com/ChristopherVR/ooxml/commit/ae061fd4add660a2c21db5f2c2ffcd15b3b5d975))
- **visio:** Put patterns, line ends and effects in Format Shape ([bae5aa8](https://github.com/ChristopherVR/ooxml/commit/bae5aa8739754f26fccd429d34224f8f68766368))
- **visio:** Drop stencil shapes as master instances in the editor ([d9f1a96](https://github.com/ChristopherVR/ooxml/commit/d9f1a96ddc6e13daf9853c43e7475116a090e445))
- **visio:** Offer built-in shapes in Change Shape on stencil shapes ([e8de13f](https://github.com/ChristopherVR/ooxml/commit/e8de13f6e4500e7312ab39f2d80269a1ec052fff))

### Bug Fixes

- **pptx:** Keep the browser's image drag out of picture moves ([817a36a](https://github.com/ChristopherVR/ooxml/commit/817a36a1eae21a48e149ee5459597577dd2cdf22))
- **ui:** Keep add-in panels when only their callbacks change ([2ff04e4](https://github.com/ChristopherVR/ooxml/commit/2ff04e4295d102cb9ad1e04cfe3f330aaad26a3f))
- **ui:** Draw add-in commands in editors that never registered them ([deb027d](https://github.com/ChristopherVR/ooxml/commit/deb027dbaabbe087b488ecef68499cd0e37d7919))
- **ui:** Honour c:dispBlanksAs in combo line series ([#45](https://github.com/ChristopherVR/ooxml/issues/45)) ([00fd4b2](https://github.com/ChristopherVR/ooxml/commit/00fd4b2a5268974ee19d4f9810d421af161666aa))

### Refactor

- **visio:** Put the External Data window on the shared task pane ([a9c094a](https://github.com/ChristopherVR/ooxml/commit/a9c094a939a110ea31b0feae0b5c2ff7a2a9a58b))
- **geometry:** Move the orthogonal connector router out of pptx ([7628894](https://github.com/ChristopherVR/ooxml/commit/7628894c847d098846b7250e8cf7cfee0b8e1ecd))
- **ui:** Move Excel and Word colour pickers onto the shared grid ([ab332b6](https://github.com/ChristopherVR/ooxml/commit/ab332b6ed7a4b80bc5016de05efb3fad694c85de))

### Documentation

- **visio:** Format the Visio and ooxml-ui READMEs ([cd76ee2](https://github.com/ChristopherVR/ooxml/commit/cd76ee27e0961385949517e7e95a8405aecbcf35))
- **ui:** Describe ribbon add-in tabs as they now work in every editor ([fd056a8](https://github.com/ChristopherVR/ooxml/commit/fd056a85965d1d6a84de4ba3c880b67237632ed8))

### Testing

- **visio:** Align stencil-instance tests with the merged behaviour ([1f184fd](https://github.com/ChristopherVR/ooxml/commit/1f184fde8172da3c5fee2d204b401a11e57633b7))
- **visio:** Give the effects preset walk a longer timeout ([a0fa1f2](https://github.com/ChristopherVR/ooxml/commit/a0fa1f2b27a745b5b70acabc70d18bb3bd312b1b))
- **ui:** Run the ui suite on happy-dom by default ([f805e22](https://github.com/ChristopherVR/ooxml/commit/f805e22a365bd2b9493efc3a835e2ef1fa89681c))

### Styling

- **visio:** Format the ribbon overflow files ([235b8f0](https://github.com/ChristopherVR/ooxml/commit/235b8f0fbb7853e240c15a84d30d3416689f0c38))
- Apply oxfmt to files drifted from the formatter ([daeabf2](https://github.com/ChristopherVR/ooxml/commit/daeabf26125af866f90359ac8a5716234d5b8b36))

## [1.7.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.7.0) - 2026-10-10

### Features

- **geometry:** Share arrow-key nudges and Shift angle snapping ([a839871](https://github.com/ChristopherVR/ooxml/commit/a8398713b97bcc9a193e6bd85aba51c55066eba8))
- **visio:** Nudge with arrow keys and snap moves to the grid ([efacb75](https://github.com/ChristopherVR/ooxml/commit/efacb75c896f64a7252b9200c5e0e0d4ba0f13e9))
- **ui:** Add office-ui-task-pane and a plain tab strip variant ([8fdee5a](https://github.com/ChristopherVR/ooxml/commit/8fdee5af4f5f7f62dd77bb8974e8312d6900cb58))
- **visio:** Match Visio's task panes, status bar and in-place text ([c7c0106](https://github.com/ChristopherVR/ooxml/commit/c7c010697f52b2b2e2b580cf1914279c1e99a127))
- **visio:** Add an Office Theme for the status bar and finish its chrome ([60a8b51](https://github.com/ChristopherVR/ooxml/commit/60a8b51710832eaac158248677da4d3dadeb0052))
- **visio:** Edit formatted, field and partial text of stencil shapes ([2779656](https://github.com/ChristopherVR/ooxml/commit/27796562ef728a6be884967e9ed84323c480c16b))
- **visio:** Give the editor Visio's title bar and Office Theme: Black ([9f14793](https://github.com/ChristopherVR/ooxml/commit/9f14793f3dc0bd605454352a858b489c2c59dbbd))
- **visio:** Add Layer Properties as Visio's dialog, saved to the drawing ([8091265](https://github.com/ChristopherVR/ooxml/commit/809126590cbae8e7e41c80e3261754fe2afbb0ee))

## [1.6.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.6.2) - 2026-10-09

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## [1.6.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.6.1) - 2026-10-09

### Bug Fixes

- **pptx:** Paint the outline gradient of chart line series ([#41](https://github.com/ChristopherVR/ooxml/issues/41)) ([158d44b](https://github.com/ChristopherVR/ooxml/commit/158d44bd79582591697e9d6ef8ccc5311dc6d1f4))

## [1.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.6.0) - 2026-10-09

### Features

- **ui:** Add visio presentation mode ([7ef7441](https://github.com/ChristopherVR/ooxml/commit/7ef7441dcaa50bd97632ecf42d7a1bbaa221c0ba))
- **ui:** Enable visio freeform, pencil and arc drawing tools ([1e38ad9](https://github.com/ChristopherVR/ooxml/commit/1e38ad97a81a1f1ec023922835d501e1f7543ba2))
- **ui:** Enable the visio change shape gallery ([65b3219](https://github.com/ChristopherVR/ooxml/commit/65b32194f83ba875649898ec066b847dfc353a4a))
- **ui:** Enable format painter and change case in the visio ribbon ([55395f7](https://github.com/ChristopherVR/ooxml/commit/55395f7b12d272039a9ae5269ce65d3d13a8c75a))
- **ui:** Enable visio quick styles gallery and shadow effects ([d485ce2](https://github.com/ChristopherVR/ooxml/commit/d485ce2714783c770d81f47f55e9eb83b8f90ba7))
- **ui:** Enable visio insert pictures, link and screentip ([d4a01ad](https://github.com/ChristopherVR/ooxml/commit/d4a01ad4e80333607c04a5a97fa6b6a7926f2fa3))
- **ui:** Enable visio group and ungroup commands ([82f79c7](https://github.com/ChristopherVR/ooxml/commit/82f79c77e7e5cd2636c53dc1a4f442e80c02f154))
- **ui:** Enable the visio connector tool with glue target feedback ([2614be1](https://github.com/ChristopherVR/ooxml/commit/2614be14dfb949eebc9e11e8a04e0f4075c7186d))
- **ui:** Add visio page setup, backgrounds and page breaks ([d64077d](https://github.com/ChristopherVR/ooxml/commit/d64077d54f52b78d8b8249bc3e1011bcb0620dd7))
- **ui:** Add visio themes, variants, effects and the format shape pane ([397cf6f](https://github.com/ChristopherVR/ooxml/commit/397cf6f7b295d2703f5df3d6b976e613b46a372d))
- **ui:** Enable visio comments, reports, diagram checks and subprocess ([6d03f65](https://github.com/ChristopherVR/ooxml/commit/6d03f65939a9788b7bdf5f22c8f4b663e01c9a57))
- **ui:** Add visio more shapes, quick shapes, containers and callouts ([d115659](https://github.com/ChristopherVR/ooxml/commit/d1156595cf5685d8154718af3452d19e7e51e568))
- **ui:** Enable the visio data tab with import, linking and data graphics ([fc88751](https://github.com/ChristopherVR/ooxml/commit/fc887514fc44310a895f4183fb903be415e951a6))
- **ui:** Add visio connection points, point glue and connector styles ([53fdb86](https://github.com/ChristopherVR/ooxml/commit/53fdb86c3b6d99b73bc28346a148b1e7f40a5873))
- **ui:** Enable visio layout, layers, select by type, guides and pdf ([b278a33](https://github.com/ChristopherVR/ooxml/commit/b278a339ca7904195ad8382f8ad2c8e5f62a4833))
- **ui:** Add visio text dialog, text block tool, symbol and field ([a510104](https://github.com/ChristopherVR/ooxml/commit/a5101049323feba94182e5802b2de399ba2f019a))

### Bug Fixes

- **ui:** Make dropped visio shapes visible and never lose a drop ([fa2f36e](https://github.com/ChristopherVR/ooxml/commit/fa2f36e9c7be0c54f5649d30c0d08183b81c5ccf))
- **visio:** Show one colour rule when every value is equal ([00bc35c](https://github.com/ChristopherVR/ooxml/commit/00bc35c94781e428bca99820aca216d7030cf73e))
- **ui:** State the visio AutoConnect limit precisely ([89ac6fd](https://github.com/ChristopherVR/ooxml/commit/89ac6fd4803928ca5b4fa9e3df763d36e51970d6))
- **ui:** Keep visio text dialogs hidden and the text block overlay clear ([13eb01d](https://github.com/ChristopherVR/ooxml/commit/13eb01d47ce0147a000506cf21eeb121ac8606bd))
- **ui:** Size the symbol picker in em and drop unused visio ribbon helpers ([d4b4c11](https://github.com/ChristopherVR/ooxml/commit/d4b4c113611809a2164024f1df23964f3ea838af))
- **pptx:** Follow slide links outside slide shows ([#38](https://github.com/ChristopherVR/ooxml/issues/38)) ([f3e99cb](https://github.com/ChristopherVR/ooxml/commit/f3e99cb2d4faaefbfbb8e8c3a130744665f9be3c))

### Testing

- **ui:** Expect browser spelling and a disabled thesaurus in review ([bdc9575](https://github.com/ChristopherVR/ooxml/commit/bdc9575ec73d2b3b46f95ca2efaaee31d46abb16))

## [1.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.5.0) - 2026-10-09

### Features

- **ui:** Enable every master in the visio shapes window ([5066db2](https://github.com/ChristopherVR/ooxml/commit/5066db2dba537952c8c282161b924b3e92d2c85b))

### Bug Fixes

- **pptx:** Keep exact line spacing when a line holds a larger run ([6b8fd6a](https://github.com/ChristopherVR/ooxml/commit/6b8fd6a42dee5e0c879a98e3e945bf01d9b2bc5e))

## [1.4.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.4.1) - 2026-10-08

### Bug Fixes

- **ui:** Restore CI after the Visio find and replace changes ([f0ca01d](https://github.com/ChristopherVR/ooxml/commit/f0ca01d6618346cee74b723cfb9cb16f61eb1cb3))
- **ui:** Keep a folded ribbon group open for popup buttons ([88a9d78](https://github.com/ChristopherVR/ooxml/commit/88a9d78d5f3188c5b725d9c5baa0cd6568c954e2))
- **ui:** Restore only the device-pixel var after an export style read ([01c157b](https://github.com/ChristopherVR/ooxml/commit/01c157b4c5ec2f4248216368ef2e20b8bc282dd5))
- **xlsx:** Correct table, subtotal, save and sheet protection behavior ([#32](https://github.com/ChristopherVR/ooxml/issues/32)) ([af31ec6](https://github.com/ChristopherVR/ooxml/commit/af31ec670c229eabb8515e04bacf9f2c277ef053))
- **ui:** Keep table resize handles off merged cells ([e8b99ee](https://github.com/ChristopherVR/ooxml/commit/e8b99ee588ecd856fd4f201a96262c4a62637aef))
- **ui:** Keep morph twins from stealing in-place wheel wedges ([cd1df51](https://github.com/ChristopherVR/ooxml/commit/cd1df5149446afb5b790bc744cbf51d50e5b01b0))
- **visio:** Resolve open code-scanning alerts ([a6bb188](https://github.com/ChristopherVR/ooxml/commit/a6bb188c4da9285a09589f6716b112aee82edb8b))

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

### Testing

- **ui:** Silence expected lit and jsdom canvas warnings ([47deb26](https://github.com/ChristopherVR/ooxml/commit/47deb26c4830af0804c59a2336020e6c14170690))

## [1.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.4.0) - 2026-10-08

### Features

- **visio:** Enable source-backed formatting and shape ordering ([c39f7be](https://github.com/ChristopherVR/ooxml/commit/c39f7bee397d27afa31155a96c01997122402773))
- **visio:** Enable multi-selection formatting and arrangement ([fb70098](https://github.com/ChristopherVR/ooxml/commit/fb70098a2b100d6ec37cd091acd618c91477cfb2))
- **visio:** Enable rich formatting and source duplication ([d50e388](https://github.com/ChristopherVR/ooxml/commit/d50e388f9190809d6a860894663cdbf774ca4dca))
- **visio:** Enable source-backed shape clipboard ([2c13c1f](https://github.com/ChristopherVR/ooxml/commit/2c13c1f39998cf37e83d3949a6685f84f4b16e4d))
- **visio:** Enable new drawings and pointer selection gestures ([75ac409](https://github.com/ChristopherVR/ooxml/commit/75ac409e5e67c23203c765d3aa472b6817853090))
- **visio:** Enable resize handles and size and position pane ([7601406](https://github.com/ChristopherVR/ooxml/commit/76014064de44defbf80208cd0fab1f33ef17c00b))
- **visio:** Enable text tools and atomic drawing creation ([b8376f2](https://github.com/ChristopherVR/ooxml/commit/b8376f2575f2e594a713456d820bfb64e5048483))
- **visio:** Enable shared fill and line properties ([9ae72a0](https://github.com/ChristopherVR/ooxml/commit/9ae72a00656d2c17b8b09a51a1ac9c262654898a))

## [1.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.3.0) - 2026-10-08

### Features

- **ui:** Add pptx-ui-ribbon-actions for the powerpoint tab row ([ce94a31](https://github.com/ChristopherVR/ooxml/commit/ce94a3170ec8f09ecd83e8bf4fe424dff3225153))

### Bug Fixes

- **ui:** Keep thin pptx strokes at least one device pixel on screen ([bb1df44](https://github.com/ChristopherVR/ooxml/commit/bb1df4463b9321ce6b57214037ac938bb1d33296))
- **pptx:** Default an absent c:gapWidth to 150 ([e16f565](https://github.com/ChristopherVR/ooxml/commit/e16f5651541cfbba9b7a1280e58463c041e507d9))
- **pptx:** Clip stacked lines and areas to the plot area ([5556f31](https://github.com/ChristopherVR/ooxml/commit/5556f315cff31eb8d2be7063eb2db46ecff34d05))
- **pptx:** Label only the visible part of clipped stacked bars ([2285eef](https://github.com/ChristopherVR/ooxml/commit/2285eef11fb513364ee0994a495c41e43f33d678))
- **xlsx:** Format selection statistics with the cell number format ([649f556](https://github.com/ChristopherVR/ooxml/commit/649f5568b348083aa548b2e819da02435d68d64a))
- **docx:** Place editing mode, comments and share like office ([9fed33b](https://github.com/ChristopherVR/ooxml/commit/9fed33bda8aae213ce439510a328293d0451ce48))
- **ui:** Size the chat composer's touch layout from tokens ([42bb661](https://github.com/ChristopherVR/ooxml/commit/42bb66145246c14a93e2ade14092b78c6e2a9d09))

### Performance

- **ui:** Build the excel formula graph in idle time ([f4e3b4e](https://github.com/ChristopherVR/ooxml/commit/f4e3b4e4ce3d6d7b8221320a8165622fd4d89719))

### Refactor

- **ui:** Share the ribbon actions between excel and word ([b34bbaf](https://github.com/ChristopherVR/ooxml/commit/b34bbaf3aba00c208df249bc263a451396357d77))

## [1.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.2.0) - 2026-10-08

### Features

- **ui:** Give the visio viewer the shared theme tokens ([1121dfe](https://github.com/ChristopherVR/ooxml/commit/1121dfeea611fcf86c15b1365c7ad0beb84d51b0))
- **ui:** Extend the shared ruler for word documents ([81451bd](https://github.com/ChristopherVR/ooxml/commit/81451bd3b77b52e71d3d274f62f1887b968d2052))
- **ui:** Draw sparklines in the excel grid ([e015d99](https://github.com/ChristopherVR/ooxml/commit/e015d9963dd4429b46979873394e97022ba7d38d))
- **ui:** Extend the shared title bar and status bar tokens ([9ff07df](https://github.com/ChristopherVR/ooxml/commit/9ff07df93a6e6507cd132dafe7361a68589aec9f))
- **ui:** Add the shared comments pane ([ec3d140](https://github.com/ChristopherVR/ooxml/commit/ec3d1409316c3f3bd76c3b6fa4f4bbc4b1d22a04))
- **ui:** Share excel workbooks over the collab session ([0d274d2](https://github.com/ChristopherVR/ooxml/commit/0d274d2455cef4ef7c3f0c44578bd96a7719fbda))
- **docx:** Draw charts in the editor and print layout ([da2ed18](https://github.com/ChristopherVR/ooxml/commit/da2ed1850715cf85d01ecc4f3967aa48d9414348))
- **ui:** Add showCompatibilityToasts root option ([992adda](https://github.com/ChristopherVR/ooxml/commit/992adda597325f8d2c9ca5fcbf612ac4da1b65c0))
- **ui:** Integrate suite storage and teams host controls ([f0c9f6c](https://github.com/ChristopherVR/ooxml/commit/f0c9f6cabc3f7cbde6f8e080946375b2733856a6))
- **ui:** Add a right-aligned actions slot to the ribbon tabs ([b8943a6](https://github.com/ChristopherVR/ooxml/commit/b8943a6b49a0fa3f9f8e90debc6b8b0eb132ea8e))
- **ui:** Add a summary slot to the status bar ([c11c2b8](https://github.com/ChristopherVR/ooxml/commit/c11c2b814230a9442d970d3364a30519f6fc8ed1))

### Bug Fixes

- **ui:** Use spacing tokens in the ruler marker styles ([6acef4b](https://github.com/ChristopherVR/ooxml/commit/6acef4b098fcf47c10d8a380648caa95de1465ad))
- **xlsx:** Report sparklines as shown in the feature notes ([1db5daa](https://github.com/ChristopherVR/ooxml/commit/1db5daacd00458cdd74e6913ea5b6c627bbe6543))
- **ui:** Map chart style numbers to the palettes office uses ([c16145d](https://github.com/ChristopherVR/ooxml/commit/c16145de0c649e401259f25a0cacc06a3f5c409d))
- **ui:** Register office-ui-ruler when word creates its ruler ([e4eb8df](https://github.com/ChristopherVR/ooxml/commit/e4eb8dffbc27a863ad17da03a22572c6ffe8cd9b))
- **ui:** Publish shared visio drawings from settled state ([cb7c587](https://github.com/ChristopherVR/ooxml/commit/cb7c587e11a9d52bb86e405501f6623d2fa9eb57))
- **pptx:** Build chart style palettes over the deck theme ([7ab9527](https://github.com/ChristopherVR/ooxml/commit/7ab9527743dbb1dc1438d2d3e6e73553b13e54b7))
- **pptx:** Lay out table cell paragraphs with their own spacing ([c3db7c5](https://github.com/ChristopherVR/ooxml/commit/c3db7c57596d8f31e24b2d556ed1eea8dddad90a))
- **pptx:** Align every paragraph when a table cell is aligned ([1bb7e7d](https://github.com/ChristopherVR/ooxml/commit/1bb7e7de568a0d8a6808862fd6dc4367782e9686))
- **pptx:** Skip chart axis lines and gridlines set to no line ([c6a9fcb](https://github.com/ChristopherVR/ooxml/commit/c6a9fcb0fb07498acf6db64d9ee9417339a6f2ac))
- **pptx:** Draw value gridlines only for axes with c:majorGridlines ([3229b2e](https://github.com/ChristopherVR/ooxml/commit/3229b2e67395d24faeb803e6447fe34b07e44770))
- **pptx:** Apply a table cell run's character spacing ([6d1ee14](https://github.com/ChristopherVR/ooxml/commit/6d1ee148cbd0ea058ae032a7a3b8fe9cdace1ce1))
- **pptx:** Size stacked bars from c:gapWidth ([c951292](https://github.com/ChristopherVR/ooxml/commit/c951292dd2f4ef06b216d075554a0c5f5d6b043f))
- **pptx:** Honour c:min and c:max on stacked charts ([0895267](https://github.com/ChristopherVR/ooxml/commit/08952672d90597d8cb99102652657c6c419f5fb7))
- **xlsx:** Place editing mode, comments and share like excel ([2d161c7](https://github.com/ChristopherVR/ooxml/commit/2d161c71022b08c44d7dd7f5274acbf44f80bc1c))
- **xlsx:** Show the selection statistics before the view buttons ([c64c804](https://github.com/ChristopherVR/ooxml/commit/c64c80415c93f95db8ef48a4b4bd57f0bb0bd934))
- **ui:** List the ribbon actions and status summary slots ([65fb8af](https://github.com/ChristopherVR/ooxml/commit/65fb8afd31cd7db0bb22a4d73e90aa31b8b9c833))
- **ui:** Escape deck text in the foreignObject export document ([60ad105](https://github.com/ChristopherVR/ooxml/commit/60ad105c4753f4a9e31c24873d0da83291b037f8))

### Refactor

- **ui:** Use the Drawing type names ([02a8884](https://github.com/ChristopherVR/ooxml/commit/02a8884b2c07e8b935425fd772b9edea6c31f840))
- **ui:** Feed the xlsx office tokens from the shared bridge ([a851f3e](https://github.com/ChristopherVR/ooxml/commit/a851f3e02c6e3fc53c886aa03fdef5e62b0aca0d))
- **ui:** Drive the word ruler through office-ui-ruler ([767c09d](https://github.com/ChristopherVR/ooxml/commit/767c09dce172f384313f215bfec9ad46a4ca33f2))
- **ui:** Theme the pptx title and status bars through host tokens ([e2effd1](https://github.com/ChristopherVR/ooxml/commit/e2effd14164fa04a1476cf9f16439b5f12518528))
- **ui:** Show word comments in the shared pane ([951b4ed](https://github.com/ChristopherVR/ooxml/commit/951b4ed01e1d3f17725d15e45d960a579e7e89d1))
- **ui:** Show excel comments in the shared pane ([4a7ee1b](https://github.com/ChristopherVR/ooxml/commit/4a7ee1b40f88fd25cacb04b2ef0a6dbbc0b58138))

### Testing

- **ui:** Speed up the yjs inline comments suite ([7c8e88d](https://github.com/ChristopherVR/ooxml/commit/7c8e88d936e6539ae1e4cf12f402584bd076fcdb))
- **xlsx:** Assert one status bar and excel's tab-row actions ([91b6054](https://github.com/ChristopherVR/ooxml/commit/91b6054b0c03c637af972e9a174e5e7308f21cfb))

### Chores

- **ui:** Extend the base tsconfig in the pptx project ([581528e](https://github.com/ChristopherVR/ooxml/commit/581528e1eec8d7b8c45d79139809ba7a723939a0))
- **ui:** Regenerate the custom elements manifest for the ruler ([3a744b9](https://github.com/ChristopherVR/ooxml/commit/3a744b9fca030c32282dc1d0c9a5be39dd24dffd))
- **scripts:** Expand every wildcard in package smoke subpaths ([f240a06](https://github.com/ChristopherVR/ooxml/commit/f240a06302ce04102f52a4df11dae529009dbc3b))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.1.0) - 2026-10-08

### Features

- **ui:** Enable proven visio group rotation interactions ([be9ca55](https://github.com/ChristopherVR/ooxml/commit/be9ca5586434432f1673d42f5a0a6560ace25d83))
- **docx:** Track plain paragraph splits and joins ([0b11ba1](https://github.com/ChristopherVR/ooxml/commit/0b11ba1e3599cd082d68862010c4c2576a5b8410))

### Bug Fixes

- **docx:** Retain locked caches during field updates ([9cad7c9](https://github.com/ChristopherVR/ooxml/commit/9cad7c9b61fd76d900cef945f6766547e164e64e))
- **docx:** Repair complex field lock metadata after edits ([93b8b4e](https://github.com/ChristopherVR/ooxml/commit/93b8b4e9662c5aaf0bb7a94f599492d0441ffa93))

### Testing

- **docx:** Verify mounted tracked field replacements ([c305e45](https://github.com/ChristopherVR/ooxml/commit/c305e459da59d489e241b164cf049956e7c2c52f))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@1.0.0) - 2026-10-07

### Features

- **xlsx:** Author chart series gradients with shared controls ([a932dbd](https://github.com/ChristopherVR/ooxml/commit/a932dbd724eb5bcf6f34473bd5e6f1fc80c07849))
- **docx:** Resolve inline object revisions in shared review ([af12a22](https://github.com/ChristopherVR/ooxml/commit/af12a22e52ce03c3ce48c1b2fd41d61a0fcce91b))
- **docx:** Resolve note revisions in document history ([bc485ab](https://github.com/ChristopherVR/ooxml/commit/bc485abfb662493b8efa31fd96653b92c0954ac0))
- **xlsx:** Edit gradient stop brightness with shared color logic ([a49baa4](https://github.com/ChristopherVR/ooxml/commit/a49baa4d275182dd89bb3bd61fd5b7c9e2e719db))
- **xlsx:** Add shared gradient direction gallery ([2fff312](https://github.com/ChristopherVR/ooxml/commit/2fff3122b914ed8421128393bdcd8d937b26116c))
- **visio:** Render saved linear gradient strokes ([376480f](https://github.com/ChristopherVR/ooxml/commit/376480f823d76dcc0e9b71ace28611b6b885e755))
- **docx:** Resolve tracked paragraph marks in shared review ([f953d21](https://github.com/ChristopherVR/ooxml/commit/f953d211fd7f883e9f1eecbedb11b43d2cfc4ad3))
- **xlsx:** Drag gradient stops with shared live previews ([03fe589](https://github.com/ChristopherVR/ooxml/commit/03fe589bf37ad48b8e879a4855047586ae2e105b))
- **docx:** Resolve header and footer revisions with the body ([1229478](https://github.com/ChristopherVR/ooxml/commit/12294785b225730cee6b75816bc5fec2fd7a1e11))
- **xlsx:** Pair gradient fields with shared range controls ([dba0182](https://github.com/ChristopherVR/ooxml/commit/dba01824576dc7d329e0e4185ecf8a302f75e823))
- **docx:** Resolve inline object formatting history ([9689395](https://github.com/ChristopherVR/ooxml/commit/9689395a265d78a0f472f4880b6191130d2a9b43))
- **docx:** Project prior inline object formatting ([73691b8](https://github.com/ChristopherVR/ooxml/commit/73691b88ba16cf53b10aaaf68696ecb1c95980e8))
- **docx:** Record inline object formatting in shared history ([5de353d](https://github.com/ChristopherVR/ooxml/commit/5de353df2cf0b61b921dd2fef0141485292db1d8))
- **xlsx:** Add native named gradient preset gallery ([47ec692](https://github.com/ChristopherVR/ooxml/commit/47ec69208803df71d035ba191653134ec475ff26))
- **docx:** Preserve hard-break formatting and shared history ([a4dfbc4](https://github.com/ChristopherVR/ooxml/commit/a4dfbc40088ad577195695092149bdd195d798f2))
- **visio:** Resize native local lines through the width cell ([cdb9b5f](https://github.com/ChristopherVR/ooxml/commit/cdb9b5f5e2f393de2e740489920e70bdc820f1c2))
- **docx:** Share font commands across text and inline objects ([469c0d4](https://github.com/ChristopherVR/ooxml/commit/469c0d4b8f9d3a43dbcb2ef907ce472c2ba8e5ad))
- **xlsx:** Add rectangular gradient geometry controls ([e10c21a](https://github.com/ChristopherVR/ooxml/commit/e10c21a08bf6fdc3238504d7434db3f0b2a8b68f))
- **visio:** Drag local line endpoints on the canvas ([f68d353](https://github.com/ChristopherVR/ooxml/commit/f68d35354731f13c283032d81d2d11aec9c4b783))
- **docx:** Preserve inline hyperlinks through shared commands ([d1e7bf0](https://github.com/ChristopherVR/ooxml/commit/d1e7bf05f53d7b84fa2b7da3aec4ea006f629b8f))
- **xlsx:** Author radial and path fills on chart series ([2d3652a](https://github.com/ChristopherVR/ooxml/commit/2d3652ad10b2422302dbe457bb6899f2e56cb7bb))
- **visio:** Create straight lines through shared drawing tools ([f67c10f](https://github.com/ChristopherVR/ooxml/commit/f67c10f3032f03e4c4b5b01f1caafe0395bb7af6))
- **docx:** Share independent inline comment anchors ([4397b4b](https://github.com/ChristopherVR/ooxml/commit/4397b4bf5629a7a9da21a34b286f01d2b0853136))
- **ui:** Embed teams settings in the workspace ([6b13bae](https://github.com/ChristopherVR/ooxml/commit/6b13bae9b7fbb62a320d088a66eacc52f00a9c44))
- **ui:** Refine shared files commands and sorting ([9372d61](https://github.com/ChristopherVR/ooxml/commit/9372d61fecf826888d184a48d3f3f1c063328174))
- **xlsx:** Format chart and plot background fills ([9237f7f](https://github.com/ChristopherVR/ooxml/commit/9237f7f2ec5851088a28d2af5ca82d02806f5438))
- **visio:** Create and resize native ellipses with shared logic ([6c4d008](https://github.com/ChristopherVR/ooxml/commit/6c4d00839469577dfff31a61e27150c7bd281c53))
- **xlsx:** Format and paint chart title and legend fills ([4255dcd](https://github.com/ChristopherVR/ooxml/commit/4255dcdee2fbf4c62ea9260ffcf3a9856ebb859b))
- **visio:** Rotate local shapes through shared editing ([c68e4f4](https://github.com/ChristopherVR/ooxml/commit/c68e4f471d9fdd2927bdb47a13c95b78c0bd706e))
- **visio:** Rotate shapes with shared pointer geometry ([407fa61](https://github.com/ChristopherVR/ooxml/commit/407fa614c7f919052d44d9e31f7abfbcfe6df50d))
- **visio:** Render live rotation previews through shared svg ([87c3957](https://github.com/ChristopherVR/ooxml/commit/87c3957b5967466db44cbc95b08f5cfff21d81a7))
- **visio:** Share native quarter-turn command preparation ([c342e79](https://github.com/ChristopherVR/ooxml/commit/c342e79b82930797a7bf3a16b89a3bbe86d0128e))
- **visio:** Add shared nested quarter-turn menus ([b1c7377](https://github.com/ChristopherVR/ooxml/commit/b1c737749e22e4340ea04ddd695ada3b5049dc33))
- **visio:** Flip local shapes through shared geometry edits ([cffd415](https://github.com/ChristopherVR/ooxml/commit/cffd415a9fd7a8866f3f07dae027850dccb17c82))
- **xlsx:** Wrap chart titles with shared text flow ([14a33ed](https://github.com/ChristopherVR/ooxml/commit/14a33ed2ee58f8469e5d24a0a36923a3922f23c4))

### Bug Fixes

- **docx:** Distinguish footnotes and endnotes with shared ids ([874ff3b](https://github.com/ChristopherVR/ooxml/commit/874ff3b325a0da5675c34786f14f68cc4e31bc1d))
- **visio:** Project oblique stroke gradients in physical bounds ([2a05f37](https://github.com/ChristopherVR/ooxml/commit/2a05f37fb9e7e2e7683d0566a2831cc0122360cf))
- **visio:** Project saved oblique fills through physical bounds ([938ceae](https://github.com/ChristopherVR/ooxml/commit/938ceae0aaa1a7db365f577f36550863757f3f8f))
- **chart:** Paint native rectangular path gradients ([af81015](https://github.com/ChristopherVR/ooxml/commit/af81015bca6823dc505c26c006b2a99eb2e85c8f))
- **docx:** Locate imported inline comment anchors ([7ecbe15](https://github.com/ChristopherVR/ooxml/commit/7ecbe15f2e9b25401316579ad35e855b5add6b90))
- **visio:** Preserve drawn line endpoints beyond paper edges ([6cc3708](https://github.com/ChristopherVR/ooxml/commit/6cc37087579742326cacec4b8b670b869c1ad48f))
- **docx:** Anchor comments to complete complex fields ([065cb5c](https://github.com/ChristopherVR/ooxml/commit/065cb5c1c8554e9cdbe7b61a807d9efff1ffc2de))
- **docx:** Undo comment records with their anchors ([c736fab](https://github.com/ChristopherVR/ooxml/commit/c736fabf6d76c7e5cf87f137745e7dc1ea21bdba))
- **docx:** Retain body edits during yjs settings updates ([62d3fbd](https://github.com/ChristopherVR/ooxml/commit/62d3fbd3a5b852cb78619f3006ec982e5bbc4230))
- **docx:** Retain complete simple field comment scopes ([30500ef](https://github.com/ChristopherVR/ooxml/commit/30500efc737d69b920530893bf4de24b629b1dd9))
- **visio:** Preserve drawn boxes beyond scaled paper edges ([0048416](https://github.com/ChristopherVR/ooxml/commit/0048416f0cca19f7d3888e4c71c1d1d09583792e))
- **docx:** Retain adjacent simple field boundaries ([9303f68](https://github.com/ChristopherVR/ooxml/commit/9303f68aab372012d833669aa4eb45ceffba7726))
- **docx:** Copy field results as formatted literal text ([797db1b](https://github.com/ChristopherVR/ooxml/commit/797db1ba8f8942b2b364833e6a065504486ceaac))
- **docx:** Retain simple fields during result replacements ([908c45e](https://github.com/ChristopherVR/ooxml/commit/908c45ea9da8bf6ca75b7a4faefab87db44711d4))
- **docx:** Preserve empty fields through deletion and review ([e3c06a7](https://github.com/ChristopherVR/ooxml/commit/e3c06a73632b21645ff6553ed85d5426bfe49e53))
- **docx:** Share field-aware deletion across clipboard paths ([408ecb8](https://github.com/ChristopherVR/ooxml/commit/408ecb8a12549b953777c3037074280110ddd017))
- **xlsx:** Measure chart title and legend text in the browser ([e74cf9e](https://github.com/ChristopherVR/ooxml/commit/e74cf9e85fec51ee29fee5c2ad9f892c38a1c12f))
- **xlsx:** Align chart title lines with measured font boxes ([8499a58](https://github.com/ChristopherVR/ooxml/commit/8499a584ea11a03cb25517f3d6dad3b16104b375))
- **ui:** Preserve feedback and history for blocked visio flips ([9feb751](https://github.com/ChristopherVR/ooxml/commit/9feb75113d0e0d3f01bf509bbe55d08ee46af7ce))
- **docx:** Remove incomplete field structure from clipboard ([1539bbc](https://github.com/ChristopherVR/ooxml/commit/1539bbc9a136a904aad1e3b73f59834ac195ccd8))
- **ui:** Avoid viewport scans when clearing visio handles ([091cc4f](https://github.com/ChristopherVR/ooxml/commit/091cc4f36a833e66d59a3465d75eac1623650427))
- **ui:** Refresh the published custom element manifest ([ac3acb6](https://github.com/ChristopherVR/ooxml/commit/ac3acb6ad7af6bbbfba5536b119367ecff1e78a1))
- **pptx:** Project rectangular chart gradients in every binding ([8bc5ee3](https://github.com/ChristopherVR/ooxml/commit/8bc5ee3be9584e37a251ae72701ed9d2d392448f))
- **ci:** Clear lint errors in chart and drawing checks ([f331e67](https://github.com/ChristopherVR/ooxml/commit/f331e670ca9c52a76e048648e55d0660d925f7c4))

## [0.31.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.31.0) - 2026-10-07

### Features

- **teams:** Add embedded content previews ([c127224](https://github.com/ChristopherVR/ooxml/commit/c127224861720dd2d37e24390f940e642588b559))
- **xlsx:** Repeat copied blocks across selected paste ranges ([ac0ce3d](https://github.com/ChristopherVR/ooxml/commit/ac0ce3dd50810162fb71b285ecd1cf3e90f3138a))
- **ui:** Share the PowerPoint colour picker palette ([1c4e1a1](https://github.com/ChristopherVR/ooxml/commit/1c4e1a10f235b4d9949e598922e18fc68a757258))
- **teams:** Add shared tabs and workbook save copies ([35d637a](https://github.com/ChristopherVR/ooxml/commit/35d637acf7cebeb1453d02bd4ac792690b5ae2c0))
- **visio:** Render native-sized open arrowheads ([aca53b5](https://github.com/ChristopherVR/ooxml/commit/aca53b5ab43e771b39ed3d4f2b07adc6163a1c7f))
- **visio:** Render native curved open arrowheads ([4576cf7](https://github.com/ChristopherVR/ooxml/commit/4576cf75cef7b031ae3b44b82b0ffc85ed8a7272))
- **xlsx:** Support combined paste special options ([725d97e](https://github.com/ChristopherVR/ooxml/commit/725d97e0dd0e3759813dead0a2cd12ee0f99ee35))
- **teams:** Create and upload channel workbooks ([caa8442](https://github.com/ChristopherVR/ooxml/commit/caa8442b4f76171c9a763999e44e58dec545ebe8))
- **visio:** Match native filled arrows on straight connectors ([e4a4c57](https://github.com/ChristopherVR/ooxml/commit/e4a4c57973c025cb79764d4cb41fac152c69bc8e))
- **xlsx:** Implement native paste arithmetic operations ([bc9f9ef](https://github.com/ChristopherVR/ooxml/commit/bc9f9efa06b439fd1a216d5de9040f0c19aae24d))
- **xlsx:** Support all except borders paste ([025bf20](https://github.com/ChristopherVR/ooxml/commit/025bf20f7f85140186c766006a4db25880828fda))
- **teams:** Preview powerpoint slides inside channels ([0be35e7](https://github.com/ChristopherVR/ooxml/commit/0be35e7b096190a46ab0d9323a6d3bb9b6862e04))
- **xlsx:** Implement column widths paste ([03cfd50](https://github.com/ChristopherVR/ooxml/commit/03cfd5047fcbcf8500e0662b6bab3b432b1f5d03))
- **pptx:** Migrate shared renderer and editing into core and ui ([ba4cd0d](https://github.com/ChristopherVR/ooxml/commit/ba4cd0dc64707ed32d4386c14b80b0ad8a50300a))
- **teams:** Download local workbook copies ([62211da](https://github.com/ChristopherVR/ooxml/commit/62211da7aa37784a744d4390878b5f16d0a19112))
- **xlsx:** Paste notes and validation rules ([eb51374](https://github.com/ChristopherVR/ooxml/commit/eb51374e6d7ed77a057bcc3274d7ef0bd6b936ee))
- **teams:** Cancel file transfers and report batch progress ([9147775](https://github.com/ChristopherVR/ooxml/commit/9147775d34bc2018fd56d7c039c8762c6fa5bf7f))
- **visio:** Insert source-backed pages with native reopen checks ([d6ccef6](https://github.com/ChristopherVR/ooxml/commit/d6ccef697508fe11e8ef79bc12e1f70f8a211bd1))
- **teams:** Cancel workbook copy saves without losing edits ([f991f12](https://github.com/ChristopherVR/ooxml/commit/f991f1234cace2026d037c1203e20e10712c05ac))
- **docx:** Add shared Yjs editor collaboration ([af0ec8e](https://github.com/ChristopherVR/ooxml/commit/af0ec8e5a7d47a31433f1c33943de9ae51144b61))
- **visio:** Reorder pages with stable selection and history ([935f3cb](https://github.com/ChristopherVR/ooxml/commit/935f3cbe213c83f6ae44730dc7ff01119dfdf561))
- **docx:** Flow continuous sections on a shared page ([1909ef3](https://github.com/ChristopherVR/ooxml/commit/1909ef3b5e4935bd8ff9bbba63b83b191efff92f))
- **teams:** Render markdown tables and task lists ([d3e9827](https://github.com/ChristopherVR/ooxml/commit/d3e9827b562cb7c5824361080468484ce20b1bf4))
- **xlsx:** Preserve conditional formats through clipboard operations ([91a856a](https://github.com/ChristopherVR/ooxml/commit/91a856ae3137b960472d0df713ddb95258c85af7))
- **teams:** Add focused channel thread conversations ([e260cca](https://github.com/ChristopherVR/ooxml/commit/e260ccac17474e7325546f71f06fab74e9474e80))
- **visio:** Rename pages with references and shared dialog ([a3154d0](https://github.com/ChristopherVR/ooxml/commit/a3154d0045a39c89c9d8f77020fb06ba0e393956))
- **teams:** Add personal followed thread navigation ([e19b0c1](https://github.com/ChristopherVR/ooxml/commit/e19b0c1d466c661d28c0d59fdb5dfbe50d95fc92))
- **teams:** Track unread followed thread activity ([711fed5](https://github.com/ChristopherVR/ooxml/commit/711fed5ba0e4fb5b29112e3aa65243a3d03eca11))
- **teams:** Automatically follow authored threads ([80b1f85](https://github.com/ChristopherVR/ooxml/commit/80b1f85459a9d5b653ac587702235adef02eed86))
- **xlsx:** Render linked data-bar appearance settings ([859552f](https://github.com/ChristopherVR/ooxml/commit/859552fb68dc3be9a2f84fde914f17f8c72a7078))
- **visio:** Delete pages with native cache and background handling ([e536c35](https://github.com/ChristopherVR/ooxml/commit/e536c3553ad87ccc4e35a1d2212ab65bb22e821f))
- **docx:** Synchronize comment threads through Yjs ([817c1cc](https://github.com/ChristopherVR/ooxml/commit/817c1ccbe3116aedb41cd4b41013ed8fff3064a2))
- **pptx:** Move snapshots and cached drawing bounds into core ([f13d837](https://github.com/ChristopherVR/ooxml/commit/f13d83743ce937b5b8c5c66d0abc130d78da8dbb))
- **teams:** Preserve personal message drafts across views ([3287056](https://github.com/ChristopherVR/ooxml/commit/328705692aad1cd99a9427a0823564d703e33011))
- **xlsx:** Calculate data-bar axes and signed lengths ([42cc6e7](https://github.com/ChristopherVR/ooxml/commit/42cc6e777d77c55318105d7cf51f6b730e855e43))
- **teams:** Recover and cancel chat attachment transfers ([36332c2](https://github.com/ChristopherVR/ooxml/commit/36332c23480f95d87a093d3cde906ff405ff8019))
- **docx:** Synchronize track changes recording ([2e49393](https://github.com/ChristopherVR/ooxml/commit/2e493931f821785829127c69cd133b9b360db680))
- **visio:** Apply native layer colors and classic linear fills ([72743c2](https://github.com/ChristopherVR/ooxml/commit/72743c20f450be8c3b9c3ed9fe2741dd554a3f89))
- **docx:** Share revision recording and resolution ([26c54cd](https://github.com/ChristopherVR/ooxml/commit/26c54cd5db2c8ed69547540ed4448c979357cde4))
- **teams:** Align shared files commands and list layout ([011bf5b](https://github.com/ChristopherVR/ooxml/commit/011bf5bec452e7746e7dd4dbb0db86d9bd31f754))
- **xlsx:** Reuse shared chart gallery and edit grouping ([9dc95e9](https://github.com/ChristopherVR/ooxml/commit/9dc95e9f5797ba62f0ce0cf929dbd88a941ad2ac))
- **teams:** Organize settings and save personal appearance ([c31c15c](https://github.com/ChristopherVR/ooxml/commit/c31c15c8e7e7ef99b99ef8757dcafd491316aef2))
- **ui:** Resolve smartart theme fonts and drawing colors ([9daa589](https://github.com/ChristopherVR/ooxml/commit/9daa589d8dddba52c810efead7d7314886b580fb))
- **visio:** Render native hatch fill tiles ([1bf7486](https://github.com/ChristopherVR/ooxml/commit/1bf7486161cb286f15ec0775bae7aa1cad1124d8))
- **docx:** Restore prior run properties on revision rejection ([9427440](https://github.com/ChristopherVR/ooxml/commit/94274409912976a358179ee79a16e3c5da9f88dc))
- **chart:** Share Office chart color palettes ([bd436dd](https://github.com/ChristopherVR/ooxml/commit/bd436dd9dc392805f84a2c483c67f09e424b2e5d))
- **teams:** Add file actions and browser viewer preferences ([781872c](https://github.com/ChristopherVR/ooxml/commit/781872cab23d53a2bf11add022710d37b13c3dea))
- **docx:** Resolve imported formatting in shared review commands ([5ced959](https://github.com/ChristopherVR/ooxml/commit/5ced9591c6687df7e4c1d9773eedd855212bb328))
- **teams:** Save chat density and flush posts before navigation ([fb4284a](https://github.com/ChristopherVR/ooxml/commit/fb4284a237f54ea03c498bbff2854742122cceb5))
- **docx:** Restore prior paragraph formatting on rejection ([4612ec0](https://github.com/ChristopherVR/ooxml/commit/4612ec06e7a5e5fa70cefb51e285a8113da5e5a0))
- **teams:** Align shell menus and profile status controls ([613e54e](https://github.com/ChristopherVR/ooxml/commit/613e54e3c06e6e36c1c652acfd02c4d64756f686))
- **docx:** Resolve paragraph formatting in shared review commands ([0221310](https://github.com/ChristopherVR/ooxml/commit/02213106c5aa9d17fb48c2b247032b15b2ccaf1d))
- **visio:** Preserve page hatch orientation and tile phase ([8b13665](https://github.com/ChristopherVR/ooxml/commit/8b13665ad715f703a983938b47c6b82ba87a6aee))
- **xlsx:** Add shared Change Colors gallery and native palette edits ([c0e42f0](https://github.com/ChristopherVR/ooxml/commit/c0e42f0b611371ffef7334fa950db5ce9a98cbdd))
- **docx:** Preserve script-specific fonts through tracked edits ([a698c3a](https://github.com/ChristopherVR/ooxml/commit/a698c3a85e60d6c0bf554ec2eb10d16096caee98))
- **visio:** Render native classic radial fills ([e26a6f5](https://github.com/ChristopherVR/ooxml/commit/e26a6f5e7464f3f37b940e36137e13531be55c7b))
- **teams:** Configure shared files through an add-tab dialog ([b39ada7](https://github.com/ChristopherVR/ooxml/commit/b39ada754ac19e542c710f3e785d188cdec4b570))
- **docx:** Preserve overlapping text and formatting revisions ([88582ed](https://github.com/ChristopherVR/ooxml/commit/88582ed132cd20ef1620f1b3c59904e8650eb34c))
- **visio:** Render native classic region gradients ([92b9251](https://github.com/ChristopherVR/ooxml/commit/92b92518ac72d67630cc961882a0eb64349f4aa5))
- **teams:** Search settings and save app label preferences ([3270b52](https://github.com/ChristopherVR/ooxml/commit/3270b52ab84058a6d2e0840087eb3501038be64c))
- **docx:** Retain opaque run properties through tracked text splits ([52d5e2a](https://github.com/ChristopherVR/ooxml/commit/52d5e2a80770f4e6de7f567ab15d7a7890c719bf))
- **visio:** Render saved oblique gradients in bounding boxes ([7194727](https://github.com/ChristopherVR/ooxml/commit/71947272d8d617a2c6ee21c6e76cda9861492c91))
- **xlsx:** Render imported native chart appearance ([c14f42d](https://github.com/ChristopherVR/ooxml/commit/c14f42d99db1aad93ff522a4436aab29606d52d6))
- **docx:** Record run formatting in shared track changes ([6536756](https://github.com/ChristopherVR/ooxml/commit/653675636cfab8f927f76f9fb242d83f9eba5558))
- **docx:** Honor formatting and move recording preferences ([98cb8e3](https://github.com/ChristopherVR/ooxml/commit/98cb8e365a792352a2123a52d57ad6fc5bdcfb1a))
- **teams:** Connect shared tabs to channel conversations ([9797022](https://github.com/ChristopherVR/ooxml/commit/9797022a4fb908b6ef1df57e29901eb53d9cc5ac))
- **teams:** Refine settings switches and panel scrolling ([0e10bc5](https://github.com/ChristopherVR/ooxml/commit/0e10bc55af35b1e206d957b31838f18bcd397938))
- **docx:** Retain paragraph property bases through review edits ([af2b24c](https://github.com/ChristopherVR/ooxml/commit/af2b24cfca36df8da11de71d9b6a62f0a76ba0de))
- **teams:** Restore compact channel navigation ([861c203](https://github.com/ChristopherVR/ooxml/commit/861c203a897f3f72917918a89a2b4eef8b135845))
- **docx:** Record paragraph formatting in shared track changes ([bddb5f9](https://github.com/ChristopherVR/ooxml/commit/bddb5f9f91a9555121b5a3361721caabf10cf800))
- **xlsx:** Render imported native chart shadows ([ab83de0](https://github.com/ChristopherVR/ooxml/commit/ab83de0bc39002955cf915922d4b027f13da760e))
- **docx:** Expose shared review recording preferences in a dialog ([ea73e60](https://github.com/ChristopherVR/ooxml/commit/ea73e603f0f791c183101c2b7be8e4347c066f1a))
- **teams:** Create channels through an in-app dialog ([65bc42f](https://github.com/ChristopherVR/ooxml/commit/65bc42fc7d5214784222e20db655c8cf872f1c39))
- **xlsx:** Honor native chart gap width and overlap ([42b7d63](https://github.com/ChristopherVR/ooxml/commit/42b7d63d9bd025ca0113b6f84c9f22ea37a986ab))
- **docx:** Project prior paragraph formatting in original review ([c7260b8](https://github.com/ChristopherVR/ooxml/commit/c7260b8d5330a7e1a8e0e851674b44bce5ce85bd))
- **xlsx:** Add docked chart series spacing controls ([19ef348](https://github.com/ChristopherVR/ooxml/commit/19ef348e76d5046170edd9e50bb496dbd094207f))
- **docx:** Project prior formatting in original print layout ([582b4da](https://github.com/ChristopherVR/ooxml/commit/582b4da5fdeadcbe1d50845e20055905be6edd56))
- **xlsx:** Reuse shared range controls for chart spacing ([6d3140b](https://github.com/ChristopherVR/ooxml/commit/6d3140bdcac9578361a127c5f93e1d7185638bb5))
- **docx:** Honor review text visibility in paginated layout ([7b3fe73](https://github.com/ChristopherVR/ooxml/commit/7b3fe735269caa92795129b66caa8aa748f32c65))
- **xlsx:** Edit individual chart series fills ([94c1f0f](https://github.com/ChristopherVR/ooxml/commit/94c1f0f820d6c01b055f5c37b4c63a740fa56ea1))
- **docx:** Render prior run formatting in original review ([89c698f](https://github.com/ChristopherVR/ooxml/commit/89c698f815cbe480fa437c744d0319686ce5b1e0))
- **teams:** Render PowerPoint files with the shared DOM viewer ([eff1672](https://github.com/ChristopherVR/ooxml/commit/eff167219c4ef50759f07e02f0b920095785b693))
- **xlsx:** Add solid chart fill transparency controls ([60f5c51](https://github.com/ChristopherVR/ooxml/commit/60f5c51fcc46e8edd48d0a25c9189113d7e799b1))
- **visio:** Render saved path fills with shared gradient paints ([e0b7ee9](https://github.com/ChristopherVR/ooxml/commit/e0b7ee9953329306029d2d584562267fdd59222e))
- **docx:** Project review formatting in story editing views ([8d4c40c](https://github.com/ChristopherVR/ooxml/commit/8d4c40cc4552a1a6713037d6ca79b35c428138f5))
- **teams:** Refine settings navigation and appearance previews ([e2cd399](https://github.com/ChristopherVR/ooxml/commit/e2cd3992b7406d06ab4fa0adc20aa5d0d202a68f))

### Bug Fixes

- **xlsx:** Preserve copied blanks and repeat directional fills ([6fc8f8d](https://github.com/ChristopherVR/ooxml/commit/6fc8f8df2275936c3c62ee6ea8317a9b463e5848))
- **xlsx:** Copy cells in every ribbon fill direction ([02adfb6](https://github.com/ChristopherVR/ooxml/commit/02adfb64a83f9149c5283dc81ca3d0c75ce0386f))
- **visio:** Match native short end-arrow stems ([3219c3f](https://github.com/ChristopherVR/ooxml/commit/3219c3fb03b21c4d1645d0e02887c8d6c7e0812b))
- **visio:** Normalize drawing scales to physical page inches ([0c320b5](https://github.com/ChristopherVR/ooxml/commit/0c320b50d6edda5ce6d9a7ccba4e1c300a06833b))
- **docx:** Retain page field restarts across continuous sections ([04cce34](https://github.com/ChristopherVR/ooxml/commit/04cce34d7d6ce56deffffb05f124169e207344bb))
- **docx:** Retain concurrent overlapping comment anchors ([b131720](https://github.com/ChristopherVR/ooxml/commit/b131720f41a19dc4f4ba1145a857bd8e428b48f6))
- **ui:** Refresh Teams thread custom element manifest ([eddc15b](https://github.com/ChristopherVR/ooxml/commit/eddc15bc9ccfc281efd3e31abeb51be7936a662d))
- **docx:** Export linked moves that native word can open ([3fda2d2](https://github.com/ChristopherVR/ooxml/commit/3fda2d2d9fe2af970aff28c76fdb96089db637a2))
- **docx:** Timestamp recorded text revisions ([d657580](https://github.com/ChristopherVR/ooxml/commit/d6575805a22a7e02172aab674299e2e43953474a))
- **docx:** Preserve tracked run formatting through text edits ([ecfde3c](https://github.com/ChristopherVR/ooxml/commit/ecfde3c526ce16a41246f1ec26b22fe0a1921463))
- **docx:** Preserve paragraph revision snapshots through text edits ([30d586d](https://github.com/ChristopherVR/ooxml/commit/30d586d1893090538aa45dfe8bc1f42652638071))
- **pptx:** Preserve combo chart formatting and table subscripts ([7518420](https://github.com/ChristopherVR/ooxml/commit/75184207b96f061a557f675ef061c4c9a838393a))
- **pptx:** Apply PowerPoint table cell spacing and margins ([c5feb77](https://github.com/ChristopherVR/ooxml/commit/c5feb77c3a8e7daec77c2575092e946a3614e5ef))
- **pptx:** Resolve chart text fonts across runs and scripts ([6d00580](https://github.com/ChristopherVR/ooxml/commit/6d00580398c698e2cf4140ed2f709883b29f475d))
- **pptx:** Include bullet text runs in paragraph strut sizing ([b10c8cd](https://github.com/ChristopherVR/ooxml/commit/b10c8cd89ae3ed4ee87cf1a38f7892928858bac1))
- **pptx:** Wrap Korean text at word boundaries ([cfa6b1d](https://github.com/ChristopherVR/ooxml/commit/cfa6b1d9d27b43685b6b5e1a6ef7cf7a63a73229))
- **teams:** Declare chat density spacing tokens ([e122287](https://github.com/ChristopherVR/ooxml/commit/e122287ef9960899b72792e38f89c37c265d3244))
- **docx:** Honor style page breaks and invalidate cached markers ([64d02c8](https://github.com/ChristopherVR/ooxml/commit/64d02c86d2205380c969fc4be3d06476901f9c65))
- **visio:** Report native gradient raster fidelity gaps ([f30a2e4](https://github.com/ChristopherVR/ooxml/commit/f30a2e42f3f3fe8f815c86cd88dc7ce028ddb0ce))
- **visio:** Reproduce native opaque gradient paint ([d6930d8](https://github.com/ChristopherVR/ooxml/commit/d6930d87fe0a3d8ddced21d66094e6d63776d042))
- **visio:** Remove transparent region fill seams ([562e07e](https://github.com/ChristopherVR/ooxml/commit/562e07e8bf41be7b23e3c27e6166c98af0a2a830))
- **docx:** Preserve and filter inline object revisions ([0125d83](https://github.com/ChristopherVR/ooxml/commit/0125d83a8fe243f663c6867d76e7ffd79390d3cc))

### Refactor

- **docx:** Share comment editing commands in core ([51ad69c](https://github.com/ChristopherVR/ooxml/commit/51ad69ccd3b61e2472bbee40fda65b3eaed43e3c))
- **visio:** Use extensionless relative source imports ([2495aa9](https://github.com/ChristopherVR/ooxml/commit/2495aa92d52a0135106695a2a39a72c9f98cf177))
- **docx:** Share mark-to-run formatting conversion ([cba7fb1](https://github.com/ChristopherVR/ooxml/commit/cba7fb1ce7733a3faf6c9e463124ae69a47f6835))
- Use extensionless relative typescript imports ([ef81f39](https://github.com/ChristopherVR/ooxml/commit/ef81f394cd2f275bcc9db87830db24b0adb2ec8b))
- **docx:** Share paragraph attribute conversion ([132ce8b](https://github.com/ChristopherVR/ooxml/commit/132ce8ba237a61bbe2dbc93a0d76f80da17b0c00))
- **pptx:** Share the DOM renderer for embedded viewers ([2e3219c](https://github.com/ChristopherVR/ooxml/commit/2e3219cff733e1f49094806119f3da78a50762ee))

## [0.30.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.30.1) - 2026-10-07

## [0.30.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.30.0) - 2026-10-06

### Features

- **ui:** Add the suite tab model and a tabbed launcher shell ([8ede628](https://github.com/ChristopherVR/ooxml/commit/8ede628c6421794f93ec281eeb7f0e163d574b14))
- **ui:** Pin and drag suite tabs, and give the top bar one brand mark ([0a30d8a](https://github.com/ChristopherVR/ooxml/commit/0a30d8a4974d53eddedd896cf01b1820cee4a223))

### Bug Fixes

- **xlsx:** Harden code-scanning findings across areas ([90ee4b3](https://github.com/ChristopherVR/ooxml/commit/90ee4b33e1253f21cd0df299af468d550f85eddc))
- **xlsx:** Use Reflect for dynamic property writes and tighten sanitizers ([b663413](https://github.com/ChristopherVR/ooxml/commit/b66341384d972d73182b06c6b4c706c0462f9364))

### Dependencies

- **deps:** Align emf-converter on 4.8.21 ([3c2859d](https://github.com/ChristopherVR/ooxml/commit/3c2859d0a43e66666015debaeaa985888727db95))

## [0.29.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.29.3) - 2026-10-06

### Refactor

- Make src/core a workspace package (by @ChristopherVR) ([d5b8d6d](https://github.com/ChristopherVR/ooxml/commit/d5b8d6d91161c9ae631aa84436afb35105ac9517))

## [0.29.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.29.2) - 2026-10-05

### Refactor

- Move ooxml-ui from packages/ui to src/ui (by @ChristopherVR) ([86901f3](https://github.com/ChristopherVR/ooxml/commit/86901f3edaf0d74363d06488bd8f5724be9f807e))
- Move the library into src/core, next to src/ui (by @ChristopherVR) ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure (by @ChristopherVR) ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.29.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.29.1) - 2026-10-05

### Bug Fixes

- **ui:** Keep the popup hook on a closed gallery's placeholder (by @ChristopherVR) ([44f0060](https://github.com/ChristopherVR/ooxml/commit/44f0060aca918d9153f608e7c0d4d9fa999cbb11))
- **ui:** Repaint the select trigger when aria-label changes (by @ChristopherVR) ([50c38ff](https://github.com/ChristopherVR/ooxml/commit/50c38ff583b52991aaec9a81b3241396c714d28a))
- **ui:** Repaint the select trigger when aria-label changes (by @ChristopherVR) ([4a17d1c](https://github.com/ChristopherVR/ooxml/commit/4a17d1c2c5291f2b56057443151fba69ce072871))
- **ui:** Keep Tab inside the open File view and let F6 reach the status bar (by @ChristopherVR) ([7930c2a](https://github.com/ChristopherVR/ooxml/commit/7930c2ad2114a31652c3dc5966f56f81336e0579))

## [0.29.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.29.0) - 2026-10-05

### Features

- **ui:** Export the visio demo document from ooxml-ui/visio (by @ChristopherVR) ([0ea4098](https://github.com/ChristopherVR/ooxml/commit/0ea4098b294361f98b244788f3c833427bc3fe22))

## [0.28.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.28.0) - 2026-10-05

### Features

- **ui:** Move the xlsx editor element into ooxml-ui/xlsx (by @ChristopherVR) ([cdf0be0](https://github.com/ChristopherVR/ooxml/commit/cdf0be0f8f106deb356ddaca6a797ed0f0b1ec71))
- **ui:** Move the docx editor element into ooxml-ui/docx (by @ChristopherVR) ([680cc07](https://github.com/ChristopherVR/ooxml/commit/680cc07b085729b9d9b6e2fed65133ff275805da))
- **ui:** Move the visio viewer and the teams app into ooxml-ui (by @ChristopherVR) ([2b8605a](https://github.com/ChristopherVR/ooxml/commit/2b8605a4bcf750ecc84758091160ca9b7f490b1d))
- **xlsx:** Move the DOM-free editor logic into core (by @ChristopherVR) ([d6c6be1](https://github.com/ChristopherVR/ooxml/commit/d6c6be13ec5a1aa5826ca810a8a98b020d5c0878))
- **docx:** Move the DOM-free editor modules into core (by @ChristopherVR) ([e3bd270](https://github.com/ChristopherVR/ooxml/commit/e3bd2705501d0d30c715878ff2da87ff407fb961))
- **visio:** Move the DOM-free viewer modules into core (by @ChristopherVR) ([956cb90](https://github.com/ChristopherVR/ooxml/commit/956cb90371477a2e97e9036caf6f7b562d2a598b))

### Chores

- Lint with oxlint and pin oxfmt (by @ChristopherVR) ([e6e9aad](https://github.com/ChristopherVR/ooxml/commit/e6e9aade4bd74b3587acb80bd341b3e9505caf0b))

## [0.27.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.27.2) - 2026-10-05

## [0.27.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.27.1) - 2026-10-05

## [0.27.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.27.0) - 2026-10-05

### Features

- **ui:** Share KeyTip assignment and badges between products (by @ChristopherVR) ([5204d67](https://github.com/ChristopherVR/ooxml/commit/5204d676299c743b3aff21cceeb6ab1eb5fb0888))

## [0.26.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.26.0) - 2026-10-05

### Features

- **ui:** Publish a Custom Elements Manifest (by @ChristopherVR) ([2b00265](https://github.com/ChristopherVR/ooxml/commit/2b002654a786c6f1eb945e2ad2d19ab304d85e48))
- **ui:** Let KeyTips start and stop from code (by @ChristopherVR) ([88d8016](https://github.com/ChristopherVR/ooxml/commit/88d8016a9d7f2205471a71fb0e4465d51723f70a))
- **ui:** Tint the tab of a contextual ribbon panel (by @ChristopherVR) ([5e5d75c](https://github.com/ChristopherVR/ooxml/commit/5e5d75c115c4d626d59f046b2b3567c2345c4309))
- **ui:** Let a ribbon hide tabs by reason and expose its tab buttons (by @ChristopherVR) ([ba38042](https://github.com/ChristopherVR/ooxml/commit/ba38042dd99b26cd2c6d4f1183b3e1b5badd9789))
- **ui:** Add collapse and peek to the ribbon (by @ChristopherVR) ([4b187a0](https://github.com/ChristopherVR/ooxml/commit/4b187a06137caf760ded1c0ce86a970b2314eb50))
- **ui:** Share the ribbon overflow folding between products (by @ChristopherVR) ([c0a5aab](https://github.com/ChristopherVR/ooxml/commit/c0a5aab96b2fe533c531e3edcb4feb0ba5d9c013))

### Bug Fixes

- **ui:** Keep the manifest tooling out of the declaration build (by @ChristopherVR) ([4d2499f](https://github.com/ChristopherVR/ooxml/commit/4d2499fea7f523c507ea558bd570805555b0f283))
- **ui:** Read the Custom Elements Manifest in the package smoke test (by @ChristopherVR) ([15094ec](https://github.com/ChristopherVR/ooxml/commit/15094ecfe9fb69241d834932c2fd185668df193e))

## [0.25.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.25.0) - 2026-10-05

### Features

- **ui:** Let a ribbon panel hide its tab with data-tab-hidden (by @ChristopherVR) ([3bd9325](https://github.com/ChristopherVR/ooxml/commit/3bd9325e920e8049774a4b0894a7736ec244cdaf))
- **ui:** Add an actions slot to the title bar (by @ChristopherVR) ([65f1fea](https://github.com/ChristopherVR/ooxml/commit/65f1feac6dcb5cf7200247cfff28180c2bb36d38))

## [0.24.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.24.0) - 2026-10-04

### Features

- **ui:** Let products name the dialog close button (by @ChristopherVR) ([741bdcc](https://github.com/ChristopherVR/ooxml/commit/741bdcc2966d7ee1ad0718b96e49cacdbd44f8e5))
- **ui:** Let a status-bar button be disabled (by @ChristopherVR) ([2061ac6](https://github.com/ChristopherVR/ooxml/commit/2061ac613333a02f1eda16643c8cdb988832c29f))

## [0.23.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.23.0) - 2026-10-04

### Features

- **ui:** Show a shortcut hint on controlled context-menu rows (by @ChristopherVR) ([b3fd10a](https://github.com/ChristopherVR/ooxml/commit/b3fd10a2e1ed09addae70244c26d63dfb2fe3339))

## [0.22.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.22.0) - 2026-10-04

### Features

- **ui:** Let controlled context-menu rows carry a tooltip (by @ChristopherVR) ([c4af213](https://github.com/ChristopherVR/ooxml/commit/c4af213d78042caa52c5799be8457241e79524c4))
- **ui:** Let products name the zoom slider buttons and range (by @ChristopherVR) ([5beb442](https://github.com/ChristopherVR/ooxml/commit/5beb44256cd5129594ff4b9c82b1c9546cf9ca42))

## [0.21.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.21.0) - 2026-10-04

### Features

- **teams:** Add the team-workspace logic area (by @ChristopherVR) ([c3cca75](https://github.com/ChristopherVR/ooxml/commit/c3cca752c7833b1e1f318c0c67287c5d0bff3a9e))
- **ui:** Build every element as a Lit class with a sibling stylesheet (by @ChristopherVR) ([9e64958](https://github.com/ChristopherVR/ooxml/commit/9e6495863fbd50fb4be5e15dd1f3322bcab74e39))

### Bug Fixes

- **ui:** Keep attribute-only inputs and the popup contract of the viewers (by @ChristopherVR) ([21dcd34](https://github.com/ChristopherVR/ooxml/commit/21dcd34580d74fd1be06dc654b9a024522c7018d))
- **ui:** Give the gallery a slot for its light-DOM tiles (by @ChristopherVR) ([829270a](https://github.com/ChristopherVR/ooxml/commit/829270a845229b82f0bc00b067142de3d30d3813))
- **ui:** Render detached elements safely and keep subclass accessors out of Lit (by @ChristopherVR) ([36bcb29](https://github.com/ChristopherVR/ooxml/commit/36bcb296d565dde79a57c9cfb0d13b01cf78b7b1))

### Dependencies

- **deps:** Update all dependencies (by @ChristopherVR) ([d13b03a](https://github.com/ChristopherVR/ooxml/commit/d13b03ad9cb10c74655d699e7799d11aca4b30a0))

## [0.20.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.20.0) - 2026-10-04

### Features

- **ui:** Add the shared office-ui-gallery (by @ChristopherVR) ([720030d](https://github.com/ChristopherVR/ooxml/commit/720030d9ffbb9b0e45a6f1f96ed6739c4645dbd2))

## [0.19.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.19.0) - 2026-10-04

### Features

- **ui:** Add the shared office-ui-ribbon-section (by @ChristopherVR) ([8224de3](https://github.com/ChristopherVR/ooxml/commit/8224de33d6c945c112ce605ee6647f7e4c5dbff3))

## [0.18.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.18.2) - 2026-10-04

### Bug Fixes

- **ui:** Keep unchanged icon paths so a pressed button still clicks (by @ChristopherVR) ([aa9a613](https://github.com/ChristopherVR/ooxml/commit/aa9a6135067a7889464a0236763e9a566757187c))

## [0.18.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.18.1) - 2026-10-04

### Bug Fixes

- **ui:** Keep status bar clusters named while hidden (by @ChristopherVR) ([a60a0cf](https://github.com/ChristopherVR/ooxml/commit/a60a0cfef12cdaf131dfe19365b00583c22249d6))

## [0.18.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.18.0) - 2026-10-04

### Features

- **ui:** Add a controlled mode to office-ui-status-bar (by @ChristopherVR) ([cd96bb5](https://github.com/ChristopherVR/ooxml/commit/cd96bb55d407d2f5b8d95b80bda36ee33f68c005))

## [0.17.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.17.0) - 2026-10-04

### Features

- **ui:** Add the shared office-ui-title-bar (by @ChristopherVR) ([f0468b5](https://github.com/ChristopherVR/ooxml/commit/f0468b5e09291780b3b76246b468a86540bb8a51))

## [0.16.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.16.3) - 2026-10-04

### Bug Fixes

- **ui:** Draw a real switch knob and keep title bar search focus (by @ChristopherVR) ([1a28afd](https://github.com/ChristopherVR/ooxml/commit/1a28afd6a3dd867ed058c7141906184d2f4790bb))

## [0.16.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.16.2) - 2026-10-04

### Bug Fixes

- **ui:** Read an empty string as true in boolean properties (by @ChristopherVR) ([236cdce](https://github.com/ChristopherVR/ooxml/commit/236cdce367e911c8a826e0299e3e7048a59d8e81))

## [0.16.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.16.1) - 2026-10-04

### Bug Fixes

- **ui:** Keep internal fields out of framework property assignment (by @ChristopherVR) ([f66b5aa](https://github.com/ChristopherVR/ooxml/commit/f66b5aa7ef381144a64f67dd076df78faa83e1c8))

## [0.16.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.16.0) - 2026-10-04

### Features

- **ui:** Merge the controlled pptx context menu and load from require (by @ChristopherVR) ([213083b](https://github.com/ChristopherVR/ooxml/commit/213083bb83bff467445f920057053c606add99e7))

## [0.15.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.15.0) - 2026-10-03

### Features

- **ui:** Add ribbon command sizes, badges and group collapse (by @ChristopherVR) ([dbfb89e](https://github.com/ChristopherVR/ooxml/commit/dbfb89ecbd16b382642b349f33b2cb336c86985f))

## [0.14.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.14.0) - 2026-10-03

### Features

- **ui:** Move the declarative select from pptx-viewer (by @ChristopherVR) ([5e90b46](https://github.com/ChristopherVR/ooxml/commit/5e90b46c6a87c11e9a2485a8899b79275d2ef0a4))

## [0.13.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.13.0) - 2026-10-03

### Features

- **ui:** Make every control token based (by @ChristopherVR) ([5658df9](https://github.com/ChristopherVR/ooxml/commit/5658df98344d074ae10a247be19e296049edec3f))

## [0.12.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.12.0) - 2026-10-03

### Features

- **ui:** Expose checkbox and switch metrics as tokens (by @ChristopherVR) ([778431c](https://github.com/ChristopherVR/ooxml/commit/778431c1d9119d198cda35ecce03d3419a95ed05))

## [0.11.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.11.0) - 2026-10-03

### Features

- **ui:** Move shared office chrome controls from pptx-viewer (by @ChristopherVR) ([4fd820e](https://github.com/ChristopherVR/ooxml/commit/4fd820e06f2ca740532ea55678dd429d9584b874))

## [0.10.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.10.2) - 2026-10-03

### Bug Fixes

- **ui:** Stack backstage items in one column (by @ChristopherVR) ([23fd1cb](https://github.com/ChristopherVR/ooxml/commit/23fd1cb2e9ff3eccfc29b40c41e9599585a35f9b))

## [0.10.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.10.1) - 2026-10-03

### Bug Fixes

- **ui:** Step ribbon tabs from the focused tab (by @ChristopherVR) ([efe1a96](https://github.com/ChristopherVR/ooxml/commit/efe1a96934b699b0d37d69a2ea756ccbdbc16a26))

## [0.10.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.10.0) - 2026-10-03

### Features

- **ui:** Add office ribbon, backstage, find bar, ruler and print preview (by @ChristopherVR) ([c7d47f5](https://github.com/ChristopherVR/ooxml/commit/c7d47f545876c793d2fc965f8db252f172bc1d5c))

## [0.9.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.9.1) - 2026-10-03

## [0.9.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.9.0) - 2026-10-03

### Features

- **ui:** Add office options dialog and account profile (by @ChristopherVR) ([7c75493](https://github.com/ChristopherVR/ooxml/commit/7c75493d55743d0adbb466377f94f4115bee42c4))

## [0.8.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.8.0) - 2026-10-03

### Features

- **ui:** Add office keytips for keyboard ribbon access (by @ChristopherVR) ([29fbada](https://github.com/ChristopherVR/ooxml/commit/29fbadadf4843f3f9ccdfea944dcf9d41c19cf9c))

## [0.7.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.7.0) - 2026-10-03

### Features

- **ui:** Rank tell me results and show where commands live (by @ChristopherVR) ([1f2284e](https://github.com/ChristopherVR/ooxml/commit/1f2284eb47c579f4c417b0dda29a2055c506df99))

## [0.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.6.0) - 2026-10-03

### Features

- **ui:** Add office tell me command search (by @ChristopherVR) ([1b0db43](https://github.com/ChristopherVR/ooxml/commit/1b0db4335dcde1a74995c6ca19810db4df2337dd))

## [0.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.5.0) - 2026-10-03

### Features

- **ui:** Add context menu, menu separator and tab strip add button (by @ChristopherVR) ([69b3cb4](https://github.com/ChristopherVR/ooxml/commit/69b3cb4fd38e175ec1f34b64a3c190b755678d35))

## [0.4.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.4.1) - 2026-10-03

## [0.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.4.0) - 2026-10-03

### Features

- **ui:** Add icon-only menu buttons and keep menus on screen (by @ChristopherVR) ([410d508](https://github.com/ChristopherVR/ooxml/commit/410d50801b39e52dd90c3a4bc9158127d8e9497d))

## [0.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.3.0) - 2026-10-03

### Features

- **ui:** Add menu button, ribbon stack, launcher and office glyphs (by @ChristopherVR) ([83c8341](https://github.com/ChristopherVR/ooxml/commit/83c8341f39fdd7fdd1574954627340a5b8f3f752))

## [0.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.2.0) - 2026-10-03

### Features

- **ui:** Add zoom slider and document tab strip controls (by @ChristopherVR) ([9842596](https://github.com/ChristopherVR/ooxml/commit/98425968ba4d10ecca756e34aa8e21fb5618a0c8))

## [0.1.12](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.12) - 2026-10-03

## [0.1.11](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.11) - 2026-10-03

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([52d33d4](https://github.com/ChristopherVR/ooxml/commit/52d33d4c58f4cfac7e47df3c890168ca8542752b))

## [0.1.10](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.10) - 2026-10-03

### Chores

- **style:** Format existing docs and package configuration (by @ChristopherVR) ([09a7466](https://github.com/ChristopherVR/ooxml/commit/09a7466e8289e8f1cd0165e7830138f73b93fcf2))

## [0.1.9](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.9) - 2026-10-03

## [0.1.8](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.8) - 2026-10-03

## [0.1.7](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.7) - 2026-10-03

## [0.1.6](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.6) - 2026-10-03

## [0.1.5](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.5) - 2026-10-03

## [0.1.4](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.4) - 2026-10-03

## [0.1.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.3) - 2026-10-02

## [0.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-ui@0.1.2) - 2026-10-02

## [0.1.0](https://github.com/ChristopherVR/ooxml-core/releases/tag/@christophervr/office-ui@0.1.0) - 2026-10-01

### Features

- **ui:** Add the first office-ui controls, theme, presence and SmartArt (by @ChristopherVR) ([357f7ba](https://github.com/ChristopherVR/ooxml-core/commit/357f7bafc298d6a32727d29a82a8c6a6b441a85d))

### Bug Fixes

- **ci:** Keep the UI's core range stable so releases can install (by @ChristopherVR) ([70b3cea](https://github.com/ChristopherVR/ooxml-core/commit/70b3cea8c0715996275a775c42f1a9b8eadd72bb))

### Chores

- **ui:** Scaffold the @christophervr/office-ui workspace package (by @ChristopherVR) ([6bb6ddd](https://github.com/ChristopherVR/ooxml-core/commit/6bb6ddd0e1bd3594996810639232070239b225a4))
