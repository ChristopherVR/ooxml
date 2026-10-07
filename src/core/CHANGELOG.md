# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs releases whenever any published file changes, not only on
conventional commits.

## [0.24.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.24.0) - 2026-10-07

### Features

- **teams:** Add embedded content previews ([c127224](https://github.com/ChristopherVR/ooxml/commit/c127224861720dd2d37e24390f940e642588b559))
- **collab:** Add recovery actions and fix provider sync state ([828e686](https://github.com/ChristopherVR/ooxml/commit/828e6862c086f33a5d62025618cfd5f9d5fc3ce1))
- **xlsx:** Calculate native-compatible amorlinc depreciation ([d86855d](https://github.com/ChristopherVR/ooxml/commit/d86855dc7cfac049c5012306a32e7ed0750cc8e9))
- **xlsx:** Repeat copied blocks across selected paste ranges ([ac0ce3d](https://github.com/ChristopherVR/ooxml/commit/ac0ce3dd50810162fb71b285ecd1cf3e90f3138a))
- **teams:** Add shared tabs and workbook save copies ([35d637a](https://github.com/ChristopherVR/ooxml/commit/35d637acf7cebeb1453d02bd4ac792690b5ae2c0))
- **visio:** Render native-sized open arrowheads ([aca53b5](https://github.com/ChristopherVR/ooxml/commit/aca53b5ab43e771b39ed3d4f2b07adc6163a1c7f))
- **visio:** Render native curved open arrowheads ([4576cf7](https://github.com/ChristopherVR/ooxml/commit/4576cf75cef7b031ae3b44b82b0ffc85ed8a7272))
- **xlsx:** Support combined paste special options ([725d97e](https://github.com/ChristopherVR/ooxml/commit/725d97e0dd0e3759813dead0a2cd12ee0f99ee35))
- **xlsx:** Support whole-axis and relative r1c1 references ([30186f7](https://github.com/ChristopherVR/ooxml/commit/30186f70e030246fc8c8110581acad2fbc132f01))
- **teams:** Create and upload channel workbooks ([caa8442](https://github.com/ChristopherVR/ooxml/commit/caa8442b4f76171c9a763999e44e58dec545ebe8))
- **visio:** Match native filled arrows on straight connectors ([e4a4c57](https://github.com/ChristopherVR/ooxml/commit/e4a4c57973c025cb79764d4cb41fac152c69bc8e))
- **xlsx:** Implement native paste arithmetic operations ([bc9f9ef](https://github.com/ChristopherVR/ooxml/commit/bc9f9efa06b439fd1a216d5de9040f0c19aae24d))
- **xlsx:** Support all except borders paste ([025bf20](https://github.com/ChristopherVR/ooxml/commit/025bf20f7f85140186c766006a4db25880828fda))
- **xlsx:** Implement column widths paste ([03cfd50](https://github.com/ChristopherVR/ooxml/commit/03cfd5047fcbcf8500e0662b6bab3b432b1f5d03))
- **pptx:** Migrate shared renderer and editing into core and ui ([ba4cd0d](https://github.com/ChristopherVR/ooxml/commit/ba4cd0dc64707ed32d4386c14b80b0ad8a50300a))
- **xlsx:** Paste notes and validation rules ([eb51374](https://github.com/ChristopherVR/ooxml/commit/eb51374e6d7ed77a057bcc3274d7ef0bd6b936ee))
- **teams:** Cancel file transfers and report batch progress ([9147775](https://github.com/ChristopherVR/ooxml/commit/9147775d34bc2018fd56d7c039c8762c6fa5bf7f))
- **visio:** Insert source-backed pages with native reopen checks ([d6ccef6](https://github.com/ChristopherVR/ooxml/commit/d6ccef697508fe11e8ef79bc12e1f70f8a211bd1))
- **xlsx:** Preserve hyperlinks through clipboard operations ([1b7b5cc](https://github.com/ChristopherVR/ooxml/commit/1b7b5cc90a9b4863b5d150c8cf965f2e5c543cd6))
- **collab:** Support raw binary asset payloads ([c3e4dd3](https://github.com/ChristopherVR/ooxml/commit/c3e4dd343c65f15335dbee48d34c0c5524084b36))
- **docx:** Add shared Yjs editor collaboration ([af0ec8e](https://github.com/ChristopherVR/ooxml/commit/af0ec8e5a7d47a31433f1c33943de9ae51144b61))
- **visio:** Reorder pages with stable selection and history ([935f3cb](https://github.com/ChristopherVR/ooxml/commit/935f3cbe213c83f6ae44730dc7ff01119dfdf561))
- **docx:** Flow continuous sections on a shared page ([1909ef3](https://github.com/ChristopherVR/ooxml/commit/1909ef3b5e4935bd8ff9bbba63b83b191efff92f))
- **teams:** Render markdown tables and task lists ([d3e9827](https://github.com/ChristopherVR/ooxml/commit/d3e9827b562cb7c5824361080468484ce20b1bf4))
- **docx:** Balance columns before continuous section breaks ([b0343f4](https://github.com/ChristopherVR/ooxml/commit/b0343f441fbaa6b211b02a712ad1c81ab983ed29))
- **visio:** Refresh numeric page caches after page edits ([48ae105](https://github.com/ChristopherVR/ooxml/commit/48ae10578221707cefcf9e42c1b0596b500a424a))
- **xlsx:** Preserve conditional formats through clipboard operations ([91a856a](https://github.com/ChristopherVR/ooxml/commit/91a856ae3137b960472d0df713ddb95258c85af7))
- **teams:** Add focused channel thread conversations ([e260cca](https://github.com/ChristopherVR/ooxml/commit/e260ccac17474e7325546f71f06fab74e9474e80))
- **visio:** Rename pages with references and shared dialog ([a3154d0](https://github.com/ChristopherVR/ooxml/commit/a3154d0045a39c89c9d8f77020fb06ba0e393956))
- **teams:** Add personal followed thread navigation ([e19b0c1](https://github.com/ChristopherVR/ooxml/commit/e19b0c1d466c661d28c0d59fdb5dfbe50d95fc92))
- **xlsx:** Preserve linked data-bar settings during paste ([cc3bacc](https://github.com/ChristopherVR/ooxml/commit/cc3bacc1b7bb0a50bad494a5fafc63d61855a042))
- **teams:** Track unread followed thread activity ([711fed5](https://github.com/ChristopherVR/ooxml/commit/711fed5ba0e4fb5b29112e3aa65243a3d03eca11))
- **teams:** Automatically follow authored threads ([80b1f85](https://github.com/ChristopherVR/ooxml/commit/80b1f85459a9d5b653ac587702235adef02eed86))
- **xlsx:** Render linked data-bar appearance settings ([859552f](https://github.com/ChristopherVR/ooxml/commit/859552fb68dc3be9a2f84fde914f17f8c72a7078))
- **visio:** Delete pages with native cache and background handling ([e536c35](https://github.com/ChristopherVR/ooxml/commit/e536c3553ad87ccc4e35a1d2212ab65bb22e821f))
- **docx:** Synchronize comment threads through Yjs ([817c1cc](https://github.com/ChristopherVR/ooxml/commit/817c1ccbe3116aedb41cd4b41013ed8fff3064a2))
- **docx:** Balance unequal columns before continuous breaks ([af193f3](https://github.com/ChristopherVR/ooxml/commit/af193f3a3fc9650c6dff3ba5d43f4b0f027c2387))
- **pptx:** Move snapshots and cached drawing bounds into core ([f13d837](https://github.com/ChristopherVR/ooxml/commit/f13d83743ce937b5b8c5c66d0abc130d78da8dbb))
- **teams:** Preserve personal message drafts across views ([3287056](https://github.com/ChristopherVR/ooxml/commit/328705692aad1cd99a9427a0823564d703e33011))
- **xlsx:** Calculate data-bar axes and signed lengths ([42cc6e7](https://github.com/ChristopherVR/ooxml/commit/42cc6e777d77c55318105d7cf51f6b730e855e43))
- **docx:** Balance kept table rows before continuous breaks ([c9926dc](https://github.com/ChristopherVR/ooxml/commit/c9926dcb054e7e23c5b3dc2ed7866e02a326c5fc))
- **docx:** Balance split rows and repeated table headers ([d5a5cd7](https://github.com/ChristopherVR/ooxml/commit/d5a5cd73f739ea57cff5645a517725a04b17a47e))
- **teams:** Recover and cancel chat attachment transfers ([36332c2](https://github.com/ChristopherVR/ooxml/commit/36332c23480f95d87a093d3cde906ff405ff8019))
- **xlsx:** Honor advanced data-bar length percentages ([645b3b6](https://github.com/ChristopherVR/ooxml/commit/645b3b61f24d79e988e7f062c9a5bf664ba0d6da))
- **docx:** Synchronize track changes recording ([2e49393](https://github.com/ChristopherVR/ooxml/commit/2e493931f821785829127c69cd133b9b360db680))
- **visio:** Apply native layer colors and classic linear fills ([72743c2](https://github.com/ChristopherVR/ooxml/commit/72743c20f450be8c3b9c3ed9fe2741dd554a3f89))
- **docx:** Share revision recording and resolution ([26c54cd](https://github.com/ChristopherVR/ooxml/commit/26c54cd5db2c8ed69547540ed4448c979357cde4))
- **xlsx:** Reuse shared chart gallery and edit grouping ([9dc95e9](https://github.com/ChristopherVR/ooxml/commit/9dc95e9f5797ba62f0ce0cf929dbd88a941ad2ac))
- **teams:** Organize settings and save personal appearance ([c31c15c](https://github.com/ChristopherVR/ooxml/commit/c31c15c8e7e7ef99b99ef8757dcafd491316aef2))
- **visio:** Render native hatch fill tiles ([1bf7486](https://github.com/ChristopherVR/ooxml/commit/1bf7486161cb286f15ec0775bae7aa1cad1124d8))
- **docx:** Restore prior run properties on revision rejection ([9427440](https://github.com/ChristopherVR/ooxml/commit/94274409912976a358179ee79a16e3c5da9f88dc))
- **chart:** Share Office chart color palettes ([bd436dd](https://github.com/ChristopherVR/ooxml/commit/bd436dd9dc392805f84a2c483c67f09e424b2e5d))
- **docx:** Resolve imported formatting in shared review commands ([5ced959](https://github.com/ChristopherVR/ooxml/commit/5ced9591c6687df7e4c1d9773eedd855212bb328))
- **teams:** Save chat density and flush posts before navigation ([fb4284a](https://github.com/ChristopherVR/ooxml/commit/fb4284a237f54ea03c498bbff2854742122cceb5))
- **docx:** Restore prior paragraph formatting on rejection ([4612ec0](https://github.com/ChristopherVR/ooxml/commit/4612ec06e7a5e5fa70cefb51e285a8113da5e5a0))
- **docx:** Resolve paragraph formatting in shared review commands ([0221310](https://github.com/ChristopherVR/ooxml/commit/02213106c5aa9d17fb48c2b247032b15b2ccaf1d))
- **visio:** Preserve page hatch orientation and tile phase ([8b13665](https://github.com/ChristopherVR/ooxml/commit/8b13665ad715f703a983938b47c6b82ba87a6aee))
- **docx:** Resolve revisions across document stories ([3f512c0](https://github.com/ChristopherVR/ooxml/commit/3f512c04f7b8021e616a78b3662f9f518ae72691))
- **xlsx:** Add shared Change Colors gallery and native palette edits ([c0e42f0](https://github.com/ChristopherVR/ooxml/commit/c0e42f0b611371ffef7334fa950db5ce9a98cbdd))
- **docx:** Preserve script-specific fonts through tracked edits ([a698c3a](https://github.com/ChristopherVR/ooxml/commit/a698c3a85e60d6c0bf554ec2eb10d16096caee98))
- **visio:** Render native classic radial fills ([e26a6f5](https://github.com/ChristopherVR/ooxml/commit/e26a6f5e7464f3f37b940e36137e13531be55c7b))
- **docx:** Preserve overlapping text and formatting revisions ([88582ed](https://github.com/ChristopherVR/ooxml/commit/88582ed132cd20ef1620f1b3c59904e8650eb34c))
- **visio:** Render native classic region gradients ([92b9251](https://github.com/ChristopherVR/ooxml/commit/92b92518ac72d67630cc961882a0eb64349f4aa5))
- **chart:** Share native chart style definitions ([f8a6c9e](https://github.com/ChristopherVR/ooxml/commit/f8a6c9eecc12432d91de025b355f8ab1e1e1a673))
- **visio:** Normalize native saved vertical gradients ([326a7c4](https://github.com/ChristopherVR/ooxml/commit/326a7c4ab42a5740e4cd0f819233b27295254b46))
- **docx:** Retain opaque run properties through tracked text splits ([52d5e2a](https://github.com/ChristopherVR/ooxml/commit/52d5e2a80770f4e6de7f567ab15d7a7890c719bf))
- **visio:** Render saved oblique gradients in bounding boxes ([7194727](https://github.com/ChristopherVR/ooxml/commit/71947272d8d617a2c6ee21c6e76cda9861492c91))
- **xlsx:** Render imported native chart appearance ([c14f42d](https://github.com/ChristopherVR/ooxml/commit/c14f42d99db1aad93ff522a4436aab29606d52d6))
- **docx:** Record run formatting in shared track changes ([6536756](https://github.com/ChristopherVR/ooxml/commit/653675636cfab8f927f76f9fb242d83f9eba5558))
- **docx:** Honor formatting and move recording preferences ([98cb8e3](https://github.com/ChristopherVR/ooxml/commit/98cb8e365a792352a2123a52d57ad6fc5bdcfb1a))
- **teams:** Connect shared tabs to channel conversations ([9797022](https://github.com/ChristopherVR/ooxml/commit/9797022a4fb908b6ef1df57e29901eb53d9cc5ac))
- **visio:** Normalize saved radial fill presets ([9e8a56c](https://github.com/ChristopherVR/ooxml/commit/9e8a56c14b9f7f7a4fbc84448d657e75286c5bd3))
- **visio:** Reuse classic paint for saved rectangular gradients ([a85cb5c](https://github.com/ChristopherVR/ooxml/commit/a85cb5c72c45e29a027c53f0799598b562326024))
- **xlsx:** Preserve and render native chart series gradients ([50a9288](https://github.com/ChristopherVR/ooxml/commit/50a928804960c1ab373498a78ef3aab66ff5fbca))
- **docx:** Retain paragraph property bases through review edits ([af2b24c](https://github.com/ChristopherVR/ooxml/commit/af2b24cfca36df8da11de71d9b6a62f0a76ba0de))
- **docx:** Record paragraph formatting in shared track changes ([bddb5f9](https://github.com/ChristopherVR/ooxml/commit/bddb5f9f91a9555121b5a3361721caabf10cf800))
- **xlsx:** Render imported native chart shadows ([ab83de0](https://github.com/ChristopherVR/ooxml/commit/ab83de0bc39002955cf915922d4b027f13da760e))
- **docx:** Expose shared review recording preferences in a dialog ([ea73e60](https://github.com/ChristopherVR/ooxml/commit/ea73e603f0f791c183101c2b7be8e4347c066f1a))
- **teams:** Create channels through an in-app dialog ([65bc42f](https://github.com/ChristopherVR/ooxml/commit/65bc42fc7d5214784222e20db655c8cf872f1c39))
- **xlsx:** Honor native chart gap width and overlap ([42b7d63](https://github.com/ChristopherVR/ooxml/commit/42b7d63d9bd025ca0113b6f84c9f22ea37a986ab))
- **docx:** Project prior paragraph formatting in original review ([c7260b8](https://github.com/ChristopherVR/ooxml/commit/c7260b8d5330a7e1a8e0e851674b44bce5ce85bd))
- **xlsx:** Add docked chart series spacing controls ([19ef348](https://github.com/ChristopherVR/ooxml/commit/19ef348e76d5046170edd9e50bb496dbd094207f))
- **docx:** Project prior formatting in original print layout ([582b4da](https://github.com/ChristopherVR/ooxml/commit/582b4da5fdeadcbe1d50845e20055905be6edd56))
- **docx:** Honor review text visibility in paginated layout ([7b3fe73](https://github.com/ChristopherVR/ooxml/commit/7b3fe735269caa92795129b66caa8aa748f32c65))
- **xlsx:** Edit individual chart series fills ([94c1f0f](https://github.com/ChristopherVR/ooxml/commit/94c1f0f820d6c01b055f5c37b4c63a740fa56ea1))
- **xlsx:** Add solid chart fill transparency controls ([60f5c51](https://github.com/ChristopherVR/ooxml/commit/60f5c51fcc46e8edd48d0a25c9189113d7e799b1))
- **visio:** Render saved path fills with shared gradient paints ([e0b7ee9](https://github.com/ChristopherVR/ooxml/commit/e0b7ee9953329306029d2d584562267fdd59222e))

### Bug Fixes

- **visio:** Match native short connector rounding ([862a717](https://github.com/ChristopherVR/ooxml/commit/862a71733d7b2015589ea25df07c6a9a4e2afc49))
- **docx:** Honor contiguous table header rows ([2a838e8](https://github.com/ChristopherVR/ooxml/commit/2a838e8868b5479748162fc7b510e2b80418f388))
- **xlsx:** Preserve copied blanks and repeat directional fills ([6fc8f8d](https://github.com/ChristopherVR/ooxml/commit/6fc8f8df2275936c3c62ee6ea8317a9b463e5848))
- **xlsx:** Find trailing blanks in lookup reference ranges ([41dfd46](https://github.com/ChristopherVR/ooxml/commit/41dfd46cbaf6752cc13030feff0bb547e5e70f0d))
- **xlsx:** Reject scalar classic lookup arrays ([93950e2](https://github.com/ChristopherVR/ooxml/commit/93950e259c045c795836bf885a02c245d24745e6))
- **xlsx:** Match native fractional declining depreciation ([fd96914](https://github.com/ChristopherVR/ooxml/commit/fd9691400a7c69ec4be3c5e9dc53f7cde889742a))
- **xlsx:** Order mixed types in approximate lookups ([c666f2e](https://github.com/ChristopherVR/ooxml/commit/c666f2ee933be0b0be83a0ef4e3e094c1620fe61))
- **xlsx:** Match binary lookup duplicates and blank tails ([9a65bc9](https://github.com/ChristopherVR/ooxml/commit/9a65bc969e0a08a0df344936c5564ea1eb2dc027))
- **xlsx:** Distinguish omitted lookup search modes from blanks ([4756dbd](https://github.com/ChristopherVR/ooxml/commit/4756dbd1f753457b6dc030c0e05fc6d8242c744b))
- **xlsx:** Honor native reference argument omission rules ([8323315](https://github.com/ChristopherVR/ooxml/commit/8323315715c5782189c526f638b81ded288ef6c6))
- **xlsx:** Distinguish omitted offset dimensions from blank cells ([e198528](https://github.com/ChristopherVR/ooxml/commit/e1985287106ebebf08326aa39963137f023bd782))
- **xlsx:** Copy cells in every ribbon fill direction ([02adfb6](https://github.com/ChristopherVR/ooxml/commit/02adfb64a83f9149c5283dc81ca3d0c75ce0386f))
- **visio:** Match native short end-arrow stems ([3219c3f](https://github.com/ChristopherVR/ooxml/commit/3219c3fb03b21c4d1645d0e02887c8d6c7e0812b))
- **pptx:** Wrap text in static svg exports ([ece906e](https://github.com/ChristopherVR/ooxml/commit/ece906e50e8e19d41b062453ed0f898ea18abf77))
- **visio:** Normalize drawing scales to physical page inches ([0c320b5](https://github.com/ChristopherVR/ooxml/commit/0c320b50d6edda5ce6d9a7ccba4e1c300a06833b))
- **visio:** Preserve independent font lookups during edits ([6e27506](https://github.com/ChristopherVR/ooxml/commit/6e27506e999c1ad660203e19f15583b6375af4e4))
- **ci:** Resolve core PowerPoint self imports from source ([49f7f53](https://github.com/ChristopherVR/ooxml/commit/49f7f5326aa375a5078cc6ef7f600f6793e9195c))
- **pptx:** Preserve rich data label formatting when saving ([0cc1c69](https://github.com/ChristopherVR/ooxml/commit/0cc1c694ff9df5c5b88ce165f1d5602e29dd0da0))
- **docx:** Retain page field restarts across continuous sections ([04cce34](https://github.com/ChristopherVR/ooxml/commit/04cce34d7d6ce56deffffb05f124169e207344bb))
- **docx:** Retain concurrent overlapping comment anchors ([b131720](https://github.com/ChristopherVR/ooxml/commit/b131720f41a19dc4f4ba1145a857bd8e428b48f6))
- **docx:** Prevent stalled flow for oversized table lines ([63534e4](https://github.com/ChristopherVR/ooxml/commit/63534e4bb1fc41392b14451c2c44268aa40797a7))
- **docx:** Export linked moves that native word can open ([3fda2d2](https://github.com/ChristopherVR/ooxml/commit/3fda2d2d9fe2af970aff28c76fdb96089db637a2))
- **docx:** Timestamp recorded text revisions ([d657580](https://github.com/ChristopherVR/ooxml/commit/d6575805a22a7e02172aab674299e2e43953474a))
- **docx:** Preserve tracked run formatting through text edits ([ecfde3c](https://github.com/ChristopherVR/ooxml/commit/ecfde3c526ce16a41246f1ec26b22fe0a1921463))
- **ci:** Remove unsafe optional assertions from revision tests ([db594cb](https://github.com/ChristopherVR/ooxml/commit/db594cbe002c8b6a7ca1a99a7060c30fd24575a8))
- **docx:** Preserve paragraph revision snapshots through text edits ([30d586d](https://github.com/ChristopherVR/ooxml/commit/30d586d1893090538aa45dfe8bc1f42652638071))
- **docx:** Preserve pending revisions when reject all fails ([4f178ac](https://github.com/ChristopherVR/ooxml/commit/4f178ac64fd666e5b66b66c3180db513a9009a9a))
- **pptx:** Resolve chart text fonts across runs and scripts ([6d00580](https://github.com/ChristopherVR/ooxml/commit/6d00580398c698e2cf4140ed2f709883b29f475d))
- **pptx:** Include bullet text runs in paragraph strut sizing ([b10c8cd](https://github.com/ChristopherVR/ooxml/commit/b10c8cd89ae3ed4ee87cf1a38f7892928858bac1))
- **pptx:** Wrap Korean text at word boundaries ([cfa6b1d](https://github.com/ChristopherVR/ooxml/commit/cfa6b1d9d27b43685b6b5e1a6ef7cf7a63a73229))
- **docx:** Honor style page breaks and invalidate cached markers ([64d02c8](https://github.com/ChristopherVR/ooxml/commit/64d02c86d2205380c969fc4be3d06476901f9c65))
- **visio:** Report native gradient raster fidelity gaps ([f30a2e4](https://github.com/ChristopherVR/ooxml/commit/f30a2e42f3f3fe8f815c86cd88dc7ce028ddb0ce))
- **visio:** Reproduce native opaque gradient paint ([d6930d8](https://github.com/ChristopherVR/ooxml/commit/d6930d87fe0a3d8ddced21d66094e6d63776d042))
- **visio:** Remove transparent region fill seams ([562e07e](https://github.com/ChristopherVR/ooxml/commit/562e07e8bf41be7b23e3c27e6166c98af0a2a830))
- **docx:** Preserve and filter inline object revisions ([0125d83](https://github.com/ChristopherVR/ooxml/commit/0125d83a8fe243f663c6867d76e7ffd79390d3cc))

### Refactor

- **pptx:** Move chart grid editing policy into core ([0d8e23f](https://github.com/ChristopherVR/ooxml/commit/0d8e23f8cf62ed11d01d9ab0417df12d23b9b213))
- **docx:** Share comment editing commands in core ([51ad69c](https://github.com/ChristopherVR/ooxml/commit/51ad69ccd3b61e2472bbee40fda65b3eaed43e3c))
- **crypto:** Share Office GUID generation ([7d4ceeb](https://github.com/ChristopherVR/ooxml/commit/7d4ceeb07fcad2f55eef2cba73514f01af7cc2f3))
- **visio:** Use extensionless relative source imports ([2495aa9](https://github.com/ChristopherVR/ooxml/commit/2495aa92d52a0135106695a2a39a72c9f98cf177))
- **docx:** Share mark-to-run formatting conversion ([cba7fb1](https://github.com/ChristopherVR/ooxml/commit/cba7fb1ce7733a3faf6c9e463124ae69a47f6835))
- Use extensionless relative typescript imports ([ef81f39](https://github.com/ChristopherVR/ooxml/commit/ef81f394cd2f275bcc9db87830db24b0adb2ec8b))
- **docx:** Share paragraph attribute conversion ([132ce8b](https://github.com/ChristopherVR/ooxml/commit/132ce8ba237a61bbe2dbc93a0d76f80da17b0c00))

### Testing

- **xlsx:** Record native mixed-type lookup parity cases ([d4394b1](https://github.com/ChristopherVR/ooxml/commit/d4394b1488032cba2133475b553ab354f78db555))
- **xlsx:** Verify wrapped references across save and reload ([2cd5833](https://github.com/ChristopherVR/ooxml/commit/2cd583316ac5cedf438ad622e3440d5bdeeb2d43))
- **visio:** Record broader native gradient fidelity gaps ([4fd1a15](https://github.com/ChristopherVR/ooxml/commit/4fd1a151936827a10acc8b28709bb5803c9d0059))

## [0.23.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.23.0) - 2026-10-07

### Features

- **pptx:** Add PowerPoint's built-in SmartArt layout definitions ([9c6cd97](https://github.com/ChristopherVR/ooxml/commit/9c6cd97ba6276a441aadf0d527d3bdc81fabce5e))
- **pptx:** Apply built-in SmartArt layouts synchronously and save them ([4ddf31c](https://github.com/ChristopherVR/ooxml/commit/4ddf31c5f2f2a1b17e8af411ecd42bee7deaa491))

## [0.22.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.22.1) - 2026-10-06

### Bug Fixes

- **xlsx:** Remove polynomial regexes and harden local scripts ([e08d75a](https://github.com/ChristopherVR/ooxml/commit/e08d75a000f813b6d7a77a58adeccf6c0325e390))

## [0.22.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.22.0) - 2026-10-06

### Features

- **pptx:** Parse chart legend overlay from c:overlay ([dce36c6](https://github.com/ChristopherVR/ooxml/commit/dce36c6134cdaded271b73f541091b37223b6da2))
- **chart:** Move pie, radar, treemap and trendline helpers from pptx viewer ([c08a02b](https://github.com/ChristopherVR/ooxml/commit/c08a02b91b210c0f99c8c2dfbde02c2e1329fc57))
- **color:** Move text contrast and unit rgb helpers from pptx viewer ([2dcc341](https://github.com/ChristopherVR/ooxml/commit/2dcc341bd931610a9e9a8f4e8d53483cffbed8ec))
- **text:** Move tab leader and decimal tab helpers from pptx viewer ([3a5bcf0](https://github.com/ChristopherVR/ooxml/commit/3a5bcf0f4675db3d6831430c37db0c66c6814ef3))
- **geometry:** Move snap guides and align/distribute from pptx viewer ([3400e87](https://github.com/ChristopherVR/ooxml/commit/3400e8776661df9e2148e1be78819b1c12564b79))

### Bug Fixes

- **pptx:** Clear the old layout definition on a SmartArt layout switch ([086e06c](https://github.com/ChristopherVR/ooxml/commit/086e06cc28b0389da820f35b5fd4f482f2312de8))
- **xlsx:** Harden code-scanning findings across areas ([90ee4b3](https://github.com/ChristopherVR/ooxml/commit/90ee4b33e1253f21cd0df299af468d550f85eddc))
- **xlsx:** Use Reflect for dynamic property writes and tighten sanitizers ([b663413](https://github.com/ChristopherVR/ooxml/commit/b66341384d972d73182b06c6b4c706c0462f9364))

### Documentation

- **site:** Rewrite the repository README with hero, suite and quick start ([4c1f004](https://github.com/ChristopherVR/ooxml/commit/4c1f004423c92f043e8987db2c41b114d075c03c))
- Bring the pptx docs and agent guide up to date with the move ([8cc3f5e](https://github.com/ChristopherVR/ooxml/commit/8cc3f5e364232d39d2a2286e25056fd67b594108))

### Dependencies

- **deps:** Align emf-converter on 4.8.21 ([3c2859d](https://github.com/ChristopherVR/ooxml/commit/3c2859d0a43e66666015debaeaa985888727db95))

## [0.21.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.21.2) - 2026-10-06

### Refactor

- Make src/core a workspace package (by @ChristopherVR) ([d5b8d6d](https://github.com/ChristopherVR/ooxml/commit/d5b8d6d91161c9ae631aa84436afb35105ac9517))

## [0.21.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.21.1) - 2026-10-05

### Refactor

- Move ooxml-ui from packages/ui to src/ui (by @ChristopherVR) ([86901f3](https://github.com/ChristopherVR/ooxml/commit/86901f3edaf0d74363d06488bd8f5724be9f807e))
- Move the library into src/core, next to src/ui (by @ChristopherVR) ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure (by @ChristopherVR) ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.21.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.21.0) - 2026-10-05

### Features

- **chart:** Add the automatic value-axis scale (by @ChristopherVR) ([007f5b7](https://github.com/ChristopherVR/ooxml/commit/007f5b77507f99b8caf837fe0c82ea12d10465d6))
- **chart:** Add histogram and category binning (by @ChristopherVR) ([f3d564c](https://github.com/ChristopherVR/ooxml/commit/f3d564c705dc2f2ab6da4df4dcf590bd0c2df023))
- **pptx:** Add the pptx/ui subpath with the theme colour palette (by @ChristopherVR) ([0b7cde2](https://github.com/ChristopherVR/ooxml/commit/0b7cde28ecacc39bb80fbbebe1e6e9cb84e24634))
- **pptx:** Write chart spacing, axis visibility and area format (by @sedrew) ([5db45d5](https://github.com/ChristopherVR/ooxml/commit/5db45d573f28986acb8116f615ac0807a11afc91))
- **pptx:** Add chart spacing and area options to the chart SDK (by @sedrew) ([0e3e4ed](https://github.com/ChristopherVR/ooxml/commit/0e3e4ed0689e1ddfde5fdf1d0946ea60fb4bfe34))

### Bug Fixes

- **pptx:** Read series line style, keep table cell geometry fractional ([#6](https://github.com/ChristopherVR/ooxml/issues/6)) (by @IHAGI-c) ([f731f59](https://github.com/ChristopherVR/ooxml/commit/f731f59d5fc48c13e2a3032bc1566f10bae917af))
- **pptx:** Write c:scaling children in schema order (by @sedrew) ([a530d5b](https://github.com/ChristopherVR/ooxml/commit/a530d5b3f37badd4e6b12a6e7b387c23f28d84bb))

### Testing

- **pptx:** Cover chart spacing, axis visibility and area format (by @sedrew) ([5b912b7](https://github.com/ChristopherVR/ooxml/commit/5b912b7da0db97c58d56e3d9403a93a52e9408a7))

### Build & CI

- **viewers:** Make the viewers root workspaces with one lockfile (by @ChristopherVR) ([129c615](https://github.com/ChristopherVR/ooxml/commit/129c615fcf49bb15e25b136019be5425f6c36bd1))

### Chores

- **viewers:** Resolve imported viewers against this checkout (by @ChristopherVR) ([fc79891](https://github.com/ChristopherVR/ooxml/commit/fc798917be9668674973df731e56263ab063af4a))

## [0.20.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.20.0) - 2026-10-05

### Features

- **docx:** Move section edits, page setup, unit attrs and case into core (by @ChristopherVR) ([5e668b7](https://github.com/ChristopherVR/ooxml/commit/5e668b767189c06a2d739cd5809099643939ef9e))
- **xlsx:** Move the DOM-free editor logic into core (by @ChristopherVR) ([d6c6be1](https://github.com/ChristopherVR/ooxml/commit/d6c6be13ec5a1aa5826ca810a8a98b020d5c0878))
- **docx:** Move the DOM-free editor modules into core (by @ChristopherVR) ([e3bd270](https://github.com/ChristopherVR/ooxml/commit/e3bd2705501d0d30c715878ff2da87ff407fb961))
- **visio:** Move the DOM-free viewer modules into core (by @ChristopherVR) ([956cb90](https://github.com/ChristopherVR/ooxml/commit/956cb90371477a2e97e9036caf6f7b562d2a598b))

### Testing

- **xlsx:** Give the row-insert timing test headroom on shared runners (by @ChristopherVR) ([9c81002](https://github.com/ChristopherVR/ooxml/commit/9c81002188098eb8ad6230da1f80f339ea1fb8f1))

### Chores

- Lint with oxlint and pin oxfmt (by @ChristopherVR) ([e6e9aad](https://github.com/ChristopherVR/ooxml/commit/e6e9aade4bd74b3587acb80bd341b3e9505caf0b))

## [0.19.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.19.0) - 2026-10-05

### Features

- **docx:** Move column, line spacing and section helpers into core (by @ChristopherVR) ([b3beeb2](https://github.com/ChristopherVR/ooxml/commit/b3beeb2577c4cc6bcefe664f8e59499b124a8651))

## [0.18.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.18.0) - 2026-10-05

### Features

- **xlsx:** Move the formula editor text helpers into core (by @ChristopherVR) ([07e1108](https://github.com/ChristopherVR/ooxml/commit/07e110868a4287f529a87bf1e5147122f76d88df))

### Build & CI

- Add a script that reports dependency state across every repository (by @ChristopherVR) ([4b604ba](https://github.com/ChristopherVR/ooxml/commit/4b604ba58a472444eca8f77b3f0667c8a8f293b7))

## [0.17.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.17.0) - 2026-10-05

### Features

- **xlsx:** Add bond, coupon, discount-security and Treasury bill functions (by @ChristopherVR) ([761673c](https://github.com/ChristopherVR/ooxml/commit/761673c65671f8e802c69289160ace8b95bc76cb))
- **xlsx:** Add VDB, ENCODEURL, REGEXTEST, REGEXEXTRACT and REGEXREPLACE (by @ChristopherVR) ([c2c3eac](https://github.com/ChristopherVR/ooxml/commit/c2c3eacd6bd7f67a2dfa2657d5bb81b6c1a866fa))
- **xlsx:** Support iterative calculation for circular references (by @ChristopherVR) ([25c9307](https://github.com/ChristopherVR/ooxml/commit/25c930726dbbcbc8d692662347be89e786eb8452))

## [0.16.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.16.0) - 2026-10-04

### Features

- **xlsx:** Add F, beta and gamma inverses and the legacy distribution names (by @ChristopherVR) ([7218339](https://github.com/ChristopherVR/ooxml/commit/721833979bb5b9c939e08a29a1694251bb212e70))
- **xlsx:** Add Z.TEST, T.TEST, F.TEST, CHISQ.TEST and PROB (by @ChristopherVR) ([f86c3d9](https://github.com/ChristopherVR/ooxml/commit/f86c3d9cee3bf6d99c99e9d294ee19147fb9255c))
- **xlsx:** Add LINEST, LOGEST, TREND and GROWTH (by @ChristopherVR) ([d61c910](https://github.com/ChristopherVR/ooxml/commit/d61c9105d7b2ffd14426ae1e4f1a18d5da7d51f8))
- **xlsx:** Add the database functions (DSUM, DCOUNT, DGET and the rest) (by @ChristopherVR) ([8062b9d](https://github.com/ChristopherVR/ooxml/commit/8062b9d38729d74ae9e10ea279963e5d21fa03f4))

## [0.15.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.15.0) - 2026-10-04

### Features

- **teams:** Add the team-workspace logic area (by @ChristopherVR) ([c3cca75](https://github.com/ChristopherVR/ooxml/commit/c3cca752c7833b1e1f318c0c67287c5d0bff3a9e))

### Bug Fixes

- **teams:** Open files through short-lived signed links, not the token (by @ChristopherVR) ([5dcbc38](https://github.com/ChristopherVR/ooxml/commit/5dcbc38076192690018d1b37cc17fc042dfc4b6c))

### Dependencies

- **deps:** Update all dependencies (by @ChristopherVR) ([d13b03a](https://github.com/ChristopherVR/ooxml/commit/d13b03ad9cb10c74655d699e7799d11aca4b30a0))

## [0.14.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.14.1) - 2026-10-03

### Bug Fixes

- **collab:** Accept broadcast bytes cloned from another realm (by @ChristopherVR) ([03eb043](https://github.com/ChristopherVR/ooxml/commit/03eb0431c304641ba4e3f4e235625bc54c545105))

## [0.14.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.14.0) - 2026-10-03

### Features

- **collab:** Add broadcast channel transport and package adapter (by @ChristopherVR) ([9aa3f4a](https://github.com/ChristopherVR/ooxml/commit/9aa3f4ab8bbca06ede53c779e72342f3c8cb51ef))

## [0.13.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.13.0) - 2026-10-03

### Features

- **visio:** Import legacy VSD using the shared binary model (by @ChristopherVR) ([9810823](https://github.com/ChristopherVR/ooxml/commit/9810823ded5e2b1d6268914b7c9bd55a5ba62d0a))

## [0.12.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.12.0) - 2026-10-03

### Features

- **visio:** Admit proven inherited master shape moves (by @ChristopherVR) ([6cb176f](https://github.com/ChristopherVR/ooxml/commit/6cb176f4b8c58cb21b7aed30259e1dfc94371fb9))

### Performance

- **xlsx:** Reuse spill cells and index short row footprints (by @ChristopherVR) ([c3b7d37](https://github.com/ChristopherVR/ooxml/commit/c3b7d37c3cf2596a381b3b3d6ebce7a56a1e9aec))
- **xlsx:** Reuse validated full recalculation ordering (by @ChristopherVR) ([490f124](https://github.com/ChristopherVR/ooxml/commit/490f124e215853515c34d3f2504a700a62a97ee1))

## [0.11.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.11.0) - 2026-10-03

### Features

- **visio:** Move independently proven master instances (by @ChristopherVR) ([d6fc328](https://github.com/ChristopherVR/ooxml/commit/d6fc32885af6a7afe704a22337057a12c215dd3e))
- **automation:** Centralize document tools and compose MCP servers (by @ChristopherVR) ([8c586b5](https://github.com/ChristopherVR/ooxml/commit/8c586b56df6b5853cfb9dd971073a04687d985b0))

### Chores

- **style:** Format existing docs and package configuration (by @ChristopherVR) ([09a7466](https://github.com/ChristopherVR/ooxml/commit/09a7466e8289e8f1cd0165e7830138f73b93fcf2))

## [0.10.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.10.0) - 2026-10-03

### Features

- **visio:** Prove independent edits around active masters (by @ChristopherVR) ([5590c8a](https://github.com/ChristopherVR/ooxml/commit/5590c8af893e81b8a3d9f3f8874c9434821b6abd))
- **visio:** Evaluate bounded numeric formulas (by @ChristopherVR) ([2343de2](https://github.com/ChristopherVR/ooxml/commit/2343de2d82d863b7d9f765c07b8484b97a96917a))
- **crypto:** Share OOXML encryption, signatures, links and properties (by @ChristopherVR) ([050df71](https://github.com/ChristopherVR/ooxml/commit/050df71849a7ceae5bef03334c03a5bd4aeadc38))

### Bug Fixes

- **xlsx:** Fix 50+ Excel mismatches found in a full core review (by @ChristopherVR) ([4b4c7a0](https://github.com/ChristopherVR/ooxml/commit/4b4c7a006b22fe915ab9e682c37b20621c4869db))

## [0.9.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.9.0) - 2026-10-03

### Features

- Share chart text and svg geometry algorithms (by @ChristopherVR) ([bc2a7cd](https://github.com/ChristopherVR/ooxml/commit/bc2a7cdf2e5c77a2c7b1d92c63ba2b5188a1cbee))

## [0.8.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.8.0) - 2026-10-03

### Features

- **docx:** Preserve edits to cell vertical alignment (by @ChristopherVR) ([939eddf](https://github.com/ChristopherVR/ooxml/commit/939eddf3eb08a9333c399a1e46ec9c2f39150cdd))
- **docx:** Map layout fragments to source text ranges (by @ChristopherVR) ([55cbd6b](https://github.com/ChristopherVR/ooxml/commit/55cbd6bff039de1d383ef9434d3d54051e90357c))

## [0.7.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.7.0) - 2026-10-03

### Features

- **visio:** Add atomic geometry editing and bounded recalculation (by @ChristopherVR) ([dc60e01](https://github.com/ChristopherVR/ooxml/commit/dc60e01ff97ae4db5e346c9942cd4e9052ba7b9d))

## [0.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.6.0) - 2026-10-03

### Features

- **digest:** Add a shared area for synchronous digests (by @ChristopherVR) ([5094d3d](https://github.com/ChristopherVR/ooxml/commit/5094d3dd74f4313c7312867aec20ef25a4a344a4))
- **docx:** Preserve edits to individual table cell margins (by @ChristopherVR) ([c413404](https://github.com/ChristopherVR/ooxml/commit/c41340440621e94628cd26728b5947d493b2e4e7))
- **docx:** Preserve imported equations as display-only runs (by @ChristopherVR) ([7114aed](https://github.com/ChristopherVR/ooxml/commit/7114aed5e79a051dba9060129c3938b420144497))

### Bug Fixes

- **xlsx:** Verify Excel's modern sheet password hash (by @ChristopherVR) ([b38a0b7](https://github.com/ChristopherVR/ooxml/commit/b38a0b79b65f5bd00484fe9060fe1d065bb6e5ee))
- **xlsx:** Keep removed passwords removed and check every hash (by @ChristopherVR) ([5ef0c50](https://github.com/ChristopherVR/ooxml/commit/5ef0c503f1c6b48bd3ac15bb805a587a7e485d6a))

### Testing

- **pptx:** Preserve smartart across repeated save and reload (by @ChristopherVR) ([0a94952](https://github.com/ChristopherVR/ooxml/commit/0a949525b65b879dd10bad2641658e68d4c48036))

## [0.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.5.0) - 2026-10-03

### Features

- **math:** Share equation converters across office formats (by @ChristopherVR) ([2763d2a](https://github.com/ChristopherVR/ooxml/commit/2763d2afe28f0e3d7dead91d25d2011d2f618dc1))

### Bug Fixes

- **docx:** Preserve edited shading and automatic table widths (by @ChristopherVR) ([1d15587](https://github.com/ChristopherVR/ooxml/commit/1d15587d642872bda18c468d6dba13a1b2c39c33))

## [0.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.4.0) - 2026-10-02

### Features

- **visio:** Round bounded open orthogonal connectors (by @codex) ([2c6e66a](https://github.com/ChristopherVR/ooxml/commit/2c6e66afa7520e882d89498bebd117eb966476ea))
- **opc:** Add SpreadsheetML namespaces and relationship types (by @ChristopherVR) ([adc12c1](https://github.com/ChristopherVR/ooxml/commit/adc12c17c0255f56443cd5aea45e307eb6e68ae1))
- **xlsx:** Add the SpreadsheetML area (by @ChristopherVR) ([981e75b](https://github.com/ChristopherVR/ooxml/commit/981e75bf9f84dbedb8aafa99fea236c580cba9c2))

### Testing

- **xlsx:** Restore the Office user name after generating fixtures (by @ChristopherVR) ([2d9d2f9](https://github.com/ChristopherVR/ooxml/commit/2d9d2f950e3b43bad6811212850f0ecf2f0eb92b))

## [0.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.3.0) - 2026-10-02

### Features

- **visio:** Add bounded read-only VSDX scene core (by @codex) ([dee6d0d](https://github.com/ChristopherVR/ooxml/commit/dee6d0dd475d951e6fdb59994cc217674d496cd0))
- **visio:** Add bounded saves, fills and worker vector import (by @codex) ([7364222](https://github.com/ChristopherVR/ooxml/commit/7364222cc9687da5e10a5992c2bf373cb6d68d98))

## [0.2.0](https://github.com/ChristopherVR/ooxml-core/releases/tag/@christophervr/ooxml-core@0.2.0) - 2026-10-01

### Features

- **diagram:** Add a format-neutral SmartArt area (by @ChristopherVR) ([d5fc851](https://github.com/ChristopherVR/ooxml-core/commit/d5fc85134c512afc6109316b719659937c8c90da))
- **docx:** Read SmartArt drawings and preserve them on save (by @ChristopherVR) ([ab0e611](https://github.com/ChristopherVR/ooxml-core/commit/ab0e6113475918090d16249d3cbb850403d61339))
- **docx:** Add layout and load areas moved from docx-viewer (by @ChristopherVR) ([7350cc1](https://github.com/ChristopherVR/ooxml-core/commit/7350cc1ae92f5b682ebe809d173c3bd6252ea847))
- **collab:** Add the format-neutral collaboration area (by @ChristopherVR) ([f4cb051](https://github.com/ChristopherVR/ooxml-core/commit/f4cb0519a3d61cce7b6e0b829fe7ee6c064a6fd2))

### Refactor

- **pptx:** Re-import SmartArt parsers from the diagram area (by @ChristopherVR) ([f7a6f3a](https://github.com/ChristopherVR/ooxml-core/commit/f7a6f3acf856f42294ca3a9d1e051a79e22e6a0b))

### Testing

- **pptx:** Commit the e2e deck snapshot and stop reading from pptx-viewer (by @ChristopherVR) ([c3e0979](https://github.com/ChristopherVR/ooxml-core/commit/c3e09792c624155c9de416e844ba01b257218234))

### Build & CI

- Run the build steps in parallel and stop prepack rebuilding (by @ChristopherVR) ([4d25238](https://github.com/ChristopherVR/ooxml-core/commit/4d25238854ddf6169fbe4f72e535e6b78b27dfec))
- Automate releases from conventional commits with OIDC publishing (by @ChristopherVR) ([09d5a8b](https://github.com/ChristopherVR/ooxml-core/commit/09d5a8b7a09d045d525b1ce5b19b28844fcd2f6b))

### Chores

- **ui:** Scaffold the @christophervr/office-ui workspace package (by @ChristopherVR) ([6bb6ddd](https://github.com/ChristopherVR/ooxml-core/commit/6bb6ddd0e1bd3594996810639232070239b225a4))

## [0.1.0](https://github.com/ChristopherVR/ooxml-core/releases/tag/@christophervr/ooxml-core@0.1.0) - 2026-10-01

### Features

- **opc:** Add @christophervr/ooxml-opc (relationships, content types, zip helpers) (by @ChristopherVR) ([08e5427](https://github.com/ChristopherVR/ooxml-core/commit/08e5427cd24ab1ae03ba05fd0faef76a65ddeda1))
- **docx:** Move the Word core (model, parser, serializer, editing) into the docx area (by @ChristopherVR) ([b4f83e1](https://github.com/ChristopherVR/ooxml-core/commit/b4f83e123d492414258261095e72c8113459a630))
- **pptx:** Move the PowerPoint core from pptx-viewer into src/pptx (by @ChristopherVR) ([4638a2a](https://github.com/ChristopherVR/ooxml-core/commit/4638a2a991881d1e0b7fb9d89988bfd284911648))
- **pptx:** Write gradient fills for chart series and data points (by @ChristopherVR) ([3d80333](https://github.com/ChristopherVR/ooxml-core/commit/3d80333e33ed54240c488102b665fc3ef45a8377))

### Refactor

- One published package, @christophervr/ooxml-core, with an area per subpath (by @ChristopherVR) ([c8d5ef1](https://github.com/ChristopherVR/ooxml-core/commit/c8d5ef1acd2f7d545f8b5178a7101204064a0850))
- Treat docx and pptx symmetrically; publish via OIDC (by @ChristopherVR) ([1272e23](https://github.com/ChristopherVR/ooxml-core/commit/1272e2331f542bd01050303492fa9423aef359ee))

### Testing

- **pptx:** Declare group-drill.pptx in the fixture manifest; sync missing e2e decks every run (by @ChristopherVR) ([7a1b245](https://github.com/ChristopherVR/ooxml-core/commit/7a1b245f9d9bb342bcf6ea296fe3887da13c139c))

### Build & CI

- Compile ooxml-units to dist and export it for consumers (by @ChristopherVR) ([f7476df](https://github.com/ChristopherVR/ooxml-core/commit/f7476df1dc88fdaac213a44a2616afaf4fad4314))
- **pptx:** Subpath exports, relaxed tsconfig project and dual-format bundle for the pptx area (by @ChristopherVR) ([359be0b](https://github.com/ChristopherVR/ooxml-core/commit/359be0b60bb5f9b82df5d1a89fcb20d90dd4ad76))

### Chores

- Add CI, release workflow, package smoke test and community files (by @ChristopherVR) ([f1d4805](https://github.com/ChristopherVR/ooxml-core/commit/f1d48053fb86244b129de8d2bf96cc77ecf7f0eb))
