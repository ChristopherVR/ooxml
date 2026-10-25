# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs releases whenever any published file changes, not only on
conventional commits.

## [1.26.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.26.1) - 2026-10-25

### Bug Fixes

- **pptx:** Save hyperlinks to their current targets ([#204](https://github.com/ChristopherVR/ooxml/issues/204)) ([371fe79](https://github.com/ChristopherVR/ooxml/commit/371fe799d88e4158fb7ed895a1e32b0c9b67682e))

## [1.26.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.26.0) - 2026-10-25

### Features

- **xlsx:** Read a worksheet's rows without loading the workbook ([45eabde](https://github.com/ChristopherVR/ooxml/commit/45eabdebdd9e9b674007c83a9a67762d01ad672c))
- **pptx:** Let theme scheme edits name the parts they write ([51b741e](https://github.com/ChristopherVR/ooxml/commit/51b741e66bd0684edd5c44df9a5755a2885368ac))
- **pptx:** Read a theme part's format scheme with its other schemes ([868cb75](https://github.com/ChristopherVR/ooxml/commit/868cb758e84577dac0641f46eff3dad00203baa8))
- **pptx:** Put back a format scheme a theme part held ([6bcd4bb](https://github.com/ChristopherVR/ooxml/commit/6bcd4bba33314cbd9626c727df9a1b932d282230))
- **pptx:** Let the editor history forget steps it cannot restore ([a2d0a9c](https://github.com/ChristopherVR/ooxml/commit/a2d0a9ccf257647bc03bcf14f1981c78a49a5381))

### Bug Fixes

- **pptx:** Run the saves of one handler one at a time ([#184](https://github.com/ChristopherVR/ooxml/issues/184)) ([4e9353f](https://github.com/ChristopherVR/ooxml/commit/4e9353fc82d3d3ae8fd1865bc40e157465d850f5))
- **pptx:** Resolve master and layout backgrounds through their master ([c6da218](https://github.com/ChristopherVR/ooxml/commit/c6da218afe8f162ffb361d14b7dea7502c95a19e))
- **pptx:** Keep theme state on the part a switched master uses ([c806787](https://github.com/ChristopherVR/ooxml/commit/c8067874618784de1908cd81c3b5e8affda800cf))

### Performance

- **xml:** Read a fragment's child elements one at a time ([9ce77d4](https://github.com/ChristopherVR/ooxml/commit/9ce77d47f951b9e2b1183df0dac82675b6dd843c))
- **xlsx:** Read a worksheet's cell data one row at a time ([eab7423](https://github.com/ChristopherVR/ooxml/commit/eab7423eb5fba0031f1b1524d58a1a6f2a507c43))
- **xlsx:** Read shared strings without building a dom ([7cb2ec6](https://github.com/ChristopherVR/ooxml/commit/7cb2ec6fb3fa673d8f83433868f1f12afb48e4a4))

### Refactor

- **opc:** Read a guarded zip entry as a stream of chunks ([472fa92](https://github.com/ChristopherVR/ooxml/commit/472fa927be4d4dafb2b625834b30109af8c28a4a))
- **xlsx:** Share the number format view of the automation reads ([c3811da](https://github.com/ChristopherVR/ooxml/commit/c3811daeda4a9d280a47a9873fb9cd1a77fb32e2))

### Testing

- **pptx:** Give the two-themes deck parts with their own script faces ([7fe8fc7](https://github.com/ChristopherVR/ooxml/commit/7fe8fc705d55efa439df9fe61334c9c31737c9ae))
- **pptx:** Give the two-themes deck its own line and effect styles ([c42c3d5](https://github.com/ChristopherVR/ooxml/commit/c42c3d5d56cf4cb4fd3317ca369d0972f4298c63))

## [1.25.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.25.2) - 2026-10-24

### Bug Fixes

- **docx:** Paint page-view shapes without the chart frame's clip ([a07593b](https://github.com/ChristopherVR/ooxml/commit/a07593b5148ba870babdccd493dd8be826cd887a))

## [1.25.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.25.1) - 2026-10-23

### Bug Fixes

- **xlsx:** Close the open code scanning alerts in core ([0f6b643](https://github.com/ChristopherVR/ooxml/commit/0f6b64306e57958a097caaf7085d63479766acb3))

### Testing

- **pptx:** Copy fixture Buffers by value in the chart save harnesses ([#181](https://github.com/ChristopherVR/ooxml/issues/181)) ([90ce9e4](https://github.com/ChristopherVR/ooxml/commit/90ce9e4704be388ef1ee8723c0b745c3457f07e7))

## [1.25.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.25.0) - 2026-10-23

### Features

- **chart:** Model pie-of-pie split options and date axis time units ([82e3a08](https://github.com/ChristopherVR/ooxml/commit/82e3a08de62696cdd1632a07783b78019ed8e91b))
- **chart:** Add neutral chart edit operations over ChartSpace ([af3c35b](https://github.com/ChristopherVR/ooxml/commit/af3c35bc9fb61833dd94aec813d8d7a16829d733))
- **drawingml:** Add the neutral rich text body model, reader and writer ([14539db](https://github.com/ChristopherVR/ooxml/commit/14539dba2f3f8018d73dd956fb25b61d385f36a5))
- **diagram:** Share smartart node text projection and reconciliation ([a284df7](https://github.com/ChristopherVR/ooxml/commit/a284df74bc9a1dee5003be0c7b3e9df4b44b5a9c))
- **diagram:** Expose smartart node paragraphs to word and excel ([c48a282](https://github.com/ChristopherVR/ooxml/commit/c48a2821587f86bfcfd58722f06f71e00834539f))
- **chart:** Add the neutral chart data sheet model ([a681ccb](https://github.com/ChristopherVR/ooxml/commit/a681ccb4b39eef5af6a419cc340029f7bd13eb56))
- **pptx:** Export the one default undo depth ([9320c81](https://github.com/ChristopherVR/ooxml/commit/9320c81872b1508d79293b5abddc5e9abf7136db))
- **diagram:** Add the product-neutral SmartArt gallery catalogue ([219e709](https://github.com/ChristopherVR/ooxml/commit/219e709651e480cd7e6be0c26625d5933b3699be))
- **diagram:** Resolve cached smartart drawing shapes into a neutral view ([cba3609](https://github.com/ChristopherVR/ooxml/commit/cba360940c0ed00270917b2d56f5456055e10ea3))
- **chart:** Add format-target edits and typed chart edit intents ([10e1ee3](https://github.com/ChristopherVR/ooxml/commit/10e1ee3536407edc388f1ed9a4fba97632ad17b6))
- **xlsx:** Format charts in the shared chart Format pane ([f1b7e30](https://github.com/ChristopherVR/ooxml/commit/f1b7e309b76ee19828f3a83a4f8e5a43851c93a1))
- **pptx:** Save chart title and legend edits through writeChartSpace ([f29502f](https://github.com/ChristopherVR/ooxml/commit/f29502fcf5ace041780d0fbd413fb180fad38fbc))
- **pptx:** Save chart axis edits through writeChartSpace ([457eeaf](https://github.com/ChristopherVR/ooxml/commit/457eeaf1c7517380ef901d985ce96a923a008da7))
- **pptx:** Save chart series data edits through writeChartSpace ([e2c156b](https://github.com/ChristopherVR/ooxml/commit/e2c156b34224b743c2355c240e57c0973d931e17))
- **chart:** Move the embedded workbook writer into chart/workbook ([b766e1b](https://github.com/ChristopherVR/ooxml/commit/b766e1bbc67fdde7959c5c5b06bde46e41e07ccd))
- **docx:** Edit, insert and save charts through the shared chart core ([9dd3b0b](https://github.com/ChristopherVR/ooxml/commit/9dd3b0bdd3a2f6247960daf6be136ccea037b34f))
- **ui:** Insert and edit Word charts on the shared chart elements ([3227a15](https://github.com/ChristopherVR/ooxml/commit/3227a15a57ff0da6c6deb6bcd5511f734281a456))
- **chart:** Share the chart design galleries and add quick layouts ([c95c89a](https://github.com/ChristopherVR/ooxml/commit/c95c89aab3ac5cdd6cdb500e033d0b799533343e))
- **xlsx:** Add chart styles and quick layout to chart design ([a16785a](https://github.com/ChristopherVR/ooxml/commit/a16785aeb35dc33a4476c18589bb8b8287185879))
- **chart:** Apply chart data sheet edits to a chart space ([d181cda](https://github.com/ChristopherVR/ooxml/commit/d181cda21ad1c87b7b20e3a7476754b455baea43))
- **xlsx:** Read and write the chart space of an embedded workbook ([e54fd8f](https://github.com/ChristopherVR/ooxml/commit/e54fd8f90e8fe637f3d57b59a182fb88a99c7691))
- **chart:** Model, parse and write chartex parts ([07f1d0f](https://github.com/ChristopherVR/ooxml/commit/07f1d0f99ef2a3c0132feedd9c2d626045938985))
- **chart:** Add neutral chartex edit operations ([a48f991](https://github.com/ChristopherVR/ooxml/commit/a48f991b83c83ab411040f6fd7ca44bd8a24ad19))
- **pptx:** Save chartex edits through writeChartExSpace ([b4b7ca6](https://github.com/ChristopherVR/ooxml/commit/b4b7ca6ec1eb9201777419197db363ad868eaeb3))
- **diagram:** Write edits back into an existing smartart data part ([2393535](https://github.com/ChristopherVR/ooxml/commit/23935356efb880f8963bd3b03cc04a67f1bb0b27))
- **docx:** Edit, insert and save smartart ([0821917](https://github.com/ChristopherVR/ooxml/commit/0821917640adefcdc30bedd980e3c4382a390eb9))
- **docx:** Change smartart layouts and draw diagrams without a cache ([5758d49](https://github.com/ChristopherVR/ooxml/commit/5758d49d2bd307b5cbaa663392660365ea2567c6))
- **ui:** Edit and insert smartart in the word editor ([228f6b3](https://github.com/ChristopherVR/ooxml/commit/228f6b365b1343b53451670c4cc4a97f997ce179))
- **docx:** Name the chart Format Selection ribbon action ([40cc6e0](https://github.com/ChristopherVR/ooxml/commit/40cc6e07d06b179e3e57c3155555384bef79155a))
- **xlsx:** Edit, insert and save smartart ([fdfc91b](https://github.com/ChristopherVR/ooxml/commit/fdfc91b03ba214931c0123c4f4d0fafaa5a4e8e3))
- **xlsx:** Smartart ui strings, glyphs and the active smartart ([ccaa2bf](https://github.com/ChristopherVR/ooxml/commit/ccaa2bf1e3784faba3a34970c1d92db85a4a1ed2))
- **chart:** Keep unmodelled chart-space children in edited parts ([3b4e506](https://github.com/ChristopherVR/ooxml/commit/3b4e5064d421af3f6ad1a3ee37132dd93e60b79c))
- **xlsx:** Edit charts over their ChartSpace ([c5cc9ae](https://github.com/ChristopherVR/ooxml/commit/c5cc9ae412710282e3e914901822de0f983cdade))
- **ui:** Read the Excel chart Format pane from ChartSpace ([7e7717c](https://github.com/ChristopherVR/ooxml/commit/7e7717ced6b37c0a5661f7b7ca13fc3f70a08dd8))
- **chart:** Keep band formats and picture options verbatim ([502f3a4](https://github.com/ChristopherVR/ooxml/commit/502f3a4ac646fce0b79fcf4a0eaf796758eaa704))
- **pptx:** Expose smartart add shape before and above ([9427df7](https://github.com/ChristopherVR/ooxml/commit/9427df74947c45b487f944c33f33c629f5264753))
- **diagram:** Share the outline levels of a diagram's nodes ([d20e8cf](https://github.com/ChristopherVR/ooxml/commit/d20e8cf0bcfe2b9b8c5c6e06f5701adb0ef68673))

### Bug Fixes

- **chart:** Drop an unused import from the title edits ([ecb2f09](https://github.com/ChristopherVR/ooxml/commit/ecb2f0985cc5e487d51c1111859dff2a5e644bfb))
- **chart:** Keep the chart data table in line with written cells ([67387d4](https://github.com/ChristopherVR/ooxml/commit/67387d47ead5b31431143f1514b139a7728ae5be))
- **pptx:** Read bare chartex elements and named axis units ([d0db2ca](https://github.com/ChristopherVR/ooxml/commit/d0db2caf07721b154a404bbd0158867c352be3e2))
- **xlsx:** Never rewrite a chart part the model cannot carry ([ac136f7](https://github.com/ChristopherVR/ooxml/commit/ac136f7319080e129106b452504ca915310367f3))

### Other

- Neutral chart edit operations (wave6-3) ([6c283c3](https://github.com/ChristopherVR/ooxml/commit/6c283c3b863b75ae00ef407db1903c42b9785212))
- Neutral text model and shared smartart node text (wave6-4) ([10b6a35](https://github.com/ChristopherVR/ooxml/commit/10b6a359d10de68064ef8d0f525227d77631be93))
- Shared chart data sheet (wave6-5) ([b546a73](https://github.com/ChristopherVR/ooxml/commit/b546a73c0eff334308473e8d970e72cd674534e0))
- Pptx headless controllers (wave6-2) ([41dc33b](https://github.com/ChristopherVR/ooxml/commit/41dc33be23af816e33e08983f29e08621d5ae32f))
- Shared smartart picker and insert dialogs (wave6-6) ([60c4372](https://github.com/ChristopherVR/ooxml/commit/60c437219dc3e495238bb0760e17a8036e511f41))
- Smartart drawing view, edit operations and data writer (wave6-8) ([c2fe92b](https://github.com/ChristopherVR/ooxml/commit/c2fe92b05f243b5d5a9af2aed867a26957095bbd))
- Shared chart format pane (wave6-7) ([0324c80](https://github.com/ChristopherVR/ooxml/commit/0324c80633332e602ce813bc64b0e231ab6b64af))
- Pptx chart save through writeChartSpace (wave6-9) ([5ea372f](https://github.com/ChristopherVR/ooxml/commit/5ea372f859ced6b11455e8e09a0425793cf7cd67))
- Word chart editing (wave6-11) ([45ecd4b](https://github.com/ChristopherVR/ooxml/commit/45ecd4b144f4195cfa4cbe85763b87fbca162470))
- Shared chart design galleries and Visio chart editing (wave6-14) ([692e476](https://github.com/ChristopherVR/ooxml/commit/692e4763f28083d3d280b78a2e1a73e13f77499b))
- Neutral chartex model, parser, writer and edits (wave6-15) ([933b9c9](https://github.com/ChristopherVR/ooxml/commit/933b9c9bd78f3fb2825f2b5233b525791004f296))
- Bring wave6-integration into wave6-13 ([0b881ed](https://github.com/ChristopherVR/ooxml/commit/0b881ed1a85fe81295e0e68f7f92bd137e65f91b))
- Word SmartArt editing (wave6-13) ([751ba15](https://github.com/ChristopherVR/ooxml/commit/751ba157ca578e8ff6b57371a243180101ed562e))
- Word chart Format pane, chart resizing and shared pane translations (wave6-17) ([d6ac348](https://github.com/ChristopherVR/ooxml/commit/d6ac3489fd3e8ed959f4978b4de8fd4674ee71b3))
- Excel SmartArt editing on the shared diagram core (wave6-20) ([c5fdb99](https://github.com/ChristopherVR/ooxml/commit/c5fdb9919cfce94ae9daac4d1079ca235c0c5183))
- Excel chart editing over ChartSpace (wave6-18) ([ce75f7f](https://github.com/ChristopherVR/ooxml/commit/ce75f7faba40314ebc8c6de2a43c757328aa4c40))
- Pptx SmartArt text pane on the shared element (wave6-19) ([e025597](https://github.com/ChristopherVR/ooxml/commit/e025597647c9b19da2c8a308e0cc44734c569c4d))
- Bring origin/main into wave 6 ([2fedfaa](https://github.com/ChristopherVR/ooxml/commit/2fedfaae708c279044fdf86ea192315e366f80da))

### Refactor

- **xlsx:** Edit chart paint through the shared chart edit core ([95b9899](https://github.com/ChristopherVR/ooxml/commit/95b989944abb42ac2b4488dab4207eaecb0793bd))
- **drawingml:** Densify list style levels without self-assignment ([17276de](https://github.com/ChristopherVR/ooxml/commit/17276de4cc32eecc5a03de995e21aadd1d03a1af))
- **diagram:** Move smartart node operations and reflow to diagram ([bd35fff](https://github.com/ChristopherVR/ooxml/commit/bd35fff8ddee181f47358bc5a96a3969b2f330fa))
- **diagram:** Move the fabricated smartart data part writer ([d6fd6da](https://github.com/ChristopherVR/ooxml/commit/d6fd6daf55eeaefdfe3c32e772ea9b2423179998))
- **docx:** Move drawing patches and the chart run out of write-drawing ([2836055](https://github.com/ChristopherVR/ooxml/commit/2836055d0bbb5e8169e9d79cbc1d407764fca8c9))
- **chart:** Build the chartex summary from the neutral model ([78ee4b0](https://github.com/ChristopherVR/ooxml/commit/78ee4b0ab816d0338b8615834ba9bbf09c43b8cc))
- **chart:** Split the chartex axis parser and share the test part ([a514e2f](https://github.com/ChristopherVR/ooxml/commit/a514e2f8a8c655b9e3e1f67c9da23234504c4cf6))
- **diagram:** Move built-in smartart definitions to diagram/builtin ([d673145](https://github.com/ChristopherVR/ooxml/commit/d6731456c840ef85032e490f60c3492b36cd7445))
- **ui:** Ship the chart Format pane's own translations ([d6c3d65](https://github.com/ChristopherVR/ooxml/commit/d6c3d6539cf8f28ba5032959a7e2bbd0076f15ae))
- **diagram:** Share smartart node edits, tree order and inserts ([42b3ba1](https://github.com/ChristopherVR/ooxml/commit/42b3ba18f9b3369d3c69ae4615d755daaeb966fa))
- **diagram:** Share the smartart ui strings with every product ([b96c393](https://github.com/ChristopherVR/ooxml/commit/b96c393e820e88317bda0dea1ab25eb9e20b8b26))
- **xlsx:** Use the shared Format pane dictionary for chart captions ([be7364f](https://github.com/ChristopherVR/ooxml/commit/be7364fb33ba7d0a1c1c48435efe7d6afbbdd8fa))

### Testing

- **pptx:** Prove neutral chart edits load as pptx chart edits ([8ac6160](https://github.com/ChristopherVR/ooxml/commit/8ac616029886e87f6c76b829297b574f63742519))
- **pptx:** Pin smartart text parity with the neutral text model ([a7ffb36](https://github.com/ChristopherVR/ooxml/commit/a7ffb36b635521e0aaa756747f6283df3611aa32))
- **pptx:** Add the chart save harness for the neutral writer ([974c497](https://github.com/ChristopherVR/ooxml/commit/974c497dd6fc8945596bcb9b8ca04ad6dd7a1eb0))
- **docx:** Neutral author in the word smartart fixture ([e2189ef](https://github.com/ChristopherVR/ooxml/commit/e2189ef0f5cdcede5562c50d2e92a8d049a30666))
- **pptx:** Leave workbook-dependent label flags out of the save harness ([79297e5](https://github.com/ChristopherVR/ooxml/commit/79297e52145624fde9a1fb27c9c09ab7ae3323fa))
- **docx:** Check formatted and resized charts in Word acceptance ([ea06fb4](https://github.com/ChristopherVR/ooxml/commit/ea06fb4ba762089dd65c3fbdcf55bab79d9615bf))
- **xlsx:** Probe excel smartart caching and accept saved edits ([c6a6d6e](https://github.com/ChristopherVR/ooxml/commit/c6a6d6e2d9363bd57fa6f0ac444b24f4803f13e1))
- **xlsx:** Type the smartart save tests strictly ([572f7ee](https://github.com/ChristopherVR/ooxml/commit/572f7eef78ca6aba21951abc636071a788d9b593))

### Chores

- Format files the wave 6 merges left unformatted ([77bb7eb](https://github.com/ChristopherVR/ooxml/commit/77bb7eba1e3ed7984525ea862a189ac7d467f7dd))

## [1.24.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.24.3) - 2026-10-23

### Bug Fixes

- **core:** Copy Node Buffer views by value before parsing them ([#178](https://github.com/ChristopherVR/ooxml/issues/178)) ([16b3a66](https://github.com/ChristopherVR/ooxml/commit/16b3a66f5a83cf9ba198d151745c208214961021))

## [1.24.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.24.2) - 2026-10-22

### Bug Fixes

- **docx:** Follow-up fixes for content controls in table cells ([#166](https://github.com/ChristopherVR/ooxml/issues/166)) ([47b0e0f](https://github.com/ChristopherVR/ooxml/commit/47b0e0ffde6a022b40969936acb9c734de4cb8cc))
- **docx:** Keep a protected form's text field values unformatted ([#167](https://github.com/ChristopherVR/ooxml/issues/167)) ([cb21dff](https://github.com/ChristopherVR/ooxml/commit/cb21dff2988184e192b30192b807ff037a1e894e))

## [1.24.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.24.1) - 2026-10-22

### Bug Fixes

- **pptx:** Inject the .ppt svg rasteriser and move it to pptx/node ([be73b10](https://github.com/ChristopherVR/ooxml/commit/be73b10464c4629ebffa86d631b086e46a42fbe3))

### Testing

- **ci:** Check packed browser entries for node-only modules ([0420bef](https://github.com/ChristopherVR/ooxml/commit/0420bef6f98e8718a9051bc4a945980a57edf435))

## [1.24.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.24.0) - 2026-10-22

### Features

- **visio:** Resolve shared formula references by shape key ([2626b3c](https://github.com/ChristopherVR/ooxml/commit/2626b3c9dc0614787aaa7005eaa1a37494e1b759))
- **embedding:** Add the embedding area for host integration contracts ([0c4d714](https://github.com/ChristopherVR/ooxml/commit/0c4d714f4ca093fd3597e56c3c59203ae420e3a6))
- **embedding:** Add the host-owned document storage contract ([fd56d0c](https://github.com/ChristopherVR/ooxml/commit/fd56d0c6e58f06c6a153f22d7a40ff99ec593fa7))
- **visio:** Add file-command and host-owned storage to the viewer ([0c8e284](https://github.com/ChristopherVR/ooxml/commit/0c8e284015fe1cc1b895c840900482ba12f8cda4))
- **embedding:** Add the shared customization contract ([be5155d](https://github.com/ChristopherVR/ooxml/commit/be5155d20cc191cd2dd95189d53130e5979d71a7))
- **docx:** Take the shared customization contract in <docx-editor> ([65f05d3](https://github.com/ChristopherVR/ooxml/commit/65f05d3d70cc28e6633f300fb4758aca7ba34184))
- **visio:** Take the shared customization contract in <visio-viewer> ([3a6bb30](https://github.com/ChristopherVR/ooxml/commit/3a6bb305b8ac04f0942b40e8cdfbd887093c6112))
- **xlsx:** Take the shared customization contract in <xlsx-editor> ([f96a544](https://github.com/ChristopherVR/ooxml/commit/f96a5449a6a36fba59997197b484ddefb4f2d8e8))
- **embedding:** Tidy menu separators around hidden entries ([c4e6101](https://github.com/ChristopherVR/ooxml/commit/c4e6101b620e8ef9bc1c71387203232833015cf0))
- **xlsx:** Offer host menu commands on header and sheet-tab menus ([ebef1b8](https://github.com/ChristopherVR/ooxml/commit/ebef1b8a1102c568b1a7a89ab5c3390f7887aa82))
- **visio:** Accept readOnly as the inverse of canEdit ([fbc325b](https://github.com/ChristopherVR/ooxml/commit/fbc325be8c9c12a7c2e4a3affd7af2428ba4ed49))
- **embedding:** Land the host customization, read-only and storage contracts ([822a6a0](https://github.com/ChristopherVR/ooxml/commit/822a6a073840b9e9d1671476542322b28332dc42))
- **automation:** Run the document tools against open editors ([6a45ebc](https://github.com/ChristopherVR/ooxml/commit/6a45ebc3651a21739b5e18e7758395c866656198))

### Bug Fixes

- **xlsx:** Hide context-menu entries whose action is hidden ([f28dbf8](https://github.com/ChristopherVR/ooxml/commit/f28dbf85271455c1151587d475794b4cf167c2e6))
- **xlsx:** Use the shared menu separator rule and test the locks ([9f07a35](https://github.com/ChristopherVR/ooxml/commit/9f07a35cd83bd6ae64bc9300f9c7220879b29a91))

### Other

- Integrate the customization contract into host storage ([629554b](https://github.com/ChristopherVR/ooxml/commit/629554b3030728bfe16e11ac8ae5b2530bb28a17))
- Bring origin/main into the embedding work ([a3711a1](https://github.com/ChristopherVR/ooxml/commit/a3711a14049a2cd11b994e76eafec039db2dc2de))

### Performance

- **visio:** Reuse document-level data across page reads ([2f228f2](https://github.com/ChristopherVR/ooxml/commit/2f228f210bb39a02644033a30e78f09524d09948))

## [1.23.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.23.0) - 2026-10-22

### Features

- **collab:** Add a confirmable write-back scheduler ([0bc822e](https://github.com/ChristopherVR/ooxml/commit/0bc822ec620c55f7a0394ca8d685a60df31483a4))
- **visio:** Merge collaboration per cell, master and media ([46dabbc](https://github.com/ChristopherVR/ooxml/commit/46dabbc72a2ca7a0dbab0d4fda6b1bd6d4bb031e))
- **visio:** Confirm owner write-back and report its status ([ef58158](https://github.com/ChristopherVR/ooxml/commit/ef5815875434f3ce62ef0f743a16c568e8a25e45))

### Bug Fixes

- **visio:** Drop unused collab imports and type the merge spec's viewer ([74857ee](https://github.com/ChristopherVR/ooxml/commit/74857ee1d74b501d0ec0b681025e2a6879483890))

### Documentation

- Fix stale paths, counts and provenance in area docs ([95ad5db](https://github.com/ChristopherVR/ooxml/commit/95ad5db34b977df09e9a563b6ac72305cddd8ae7))
- Point script, e2e and plan references at real paths ([eec05e2](https://github.com/ChristopherVR/ooxml/commit/eec05e293e6149c0ecf2dd3c863694721949294a))

## [1.22.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.22.0) - 2026-10-22

### Features

- **automation:** Define word, excel and visio tools once in core ([bed1839](https://github.com/ChristopherVR/ooxml/commit/bed18392fec1f07fa9a0c1eabf8a7c1bb25b0387))
- **automation:** Propose xlsx edits as a reviewable diff ([0e13837](https://github.com/ChristopherVR/ooxml/commit/0e13837c6fdf0766e21aa8a31ef69737df8ab203))
- **automation:** Add xlsx_propose_edits and xlsx_apply_proposal tools ([8c7500d](https://github.com/ChristopherVR/ooxml/commit/8c7500df317f1cffa3016705369e52b0a5ad0218))
- **xlsx:** Report validation, host dialogs and state in surface props ([11ce627](https://github.com/ChristopherVR/ooxml/commit/11ce6271694e47c6fccd3d40040a44d1ae8ca4c9))
- **visio:** Let a host mark the edit history saved ([3af10a5](https://github.com/ChristopherVR/ooxml/commit/3af10a574fa5b842faaa450c6865de45e340ece4))
- **pptx:** Define the powerpoint tools as shared automation tools ([88acafd](https://github.com/ChristopherVR/ooxml/commit/88acafd82419de2cefbebd4b91b935c378f8c772))
- **collab:** Add auth token and webrtc to the embedding contract ([395e429](https://github.com/ChristopherVR/ooxml/commit/395e42961152f5725bf75bd2f0bd424f5e6f6f70))
- **visio:** Merge collaboration per shape, with cursors and write-back ([3d1564c](https://github.com/ChristopherVR/ooxml/commit/3d1564caa98d86b97bd7a384bd9d11fe533435f7))

### Bug Fixes

- **pptx:** Render flowChart connector presets as shapes, not lines ([978593a](https://github.com/ChristopherVR/ooxml/commit/978593a0d0067d97105fd98b8d640dc929ec9f2b))
- **pptx:** Paint picture fills inherited from layout placeholders ([71b8196](https://github.com/ChristopherVR/ooxml/commit/71b81964ca4b19ef9cc143fa970479614426bd9d))

## [1.21.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.21.0) - 2026-10-21

### Features

- **opc:** Name the check a zip package failed ([5125361](https://github.com/ChristopherVR/ooxml/commit/5125361f7787c1edf9863ccf85d0666a56610d1b))
- **xml:** Throw a typed error when a part breaks a parse limit ([40e73f6](https://github.com/ChristopherVR/ooxml/commit/40e73f6e0ddb045ef5a3b960e26cfc07a1fc6d45))
- **xlsx:** Let callers raise the workbook load limits ([19b3eb4](https://github.com/ChristopherVR/ooxml/commit/19b3eb4cb0cf0974529a67ea8e06b459a9dc7cf7))
- **xlsx:** Add csv import options and report lossy conversions ([29e1bb4](https://github.com/ChristopherVR/ooxml/commit/29e1bb4bbedd47e1918248e3f656738166d12636))
- **docx:** Remove temporary content controls after an edit ([d681196](https://github.com/ChristopherVR/ooxml/commit/d681196c9fd1aac6a58e053a5310476e4ab71c9a))
- **docx:** Restore content control placeholders from the glossary ([b40f123](https://github.com/ChristopherVR/ooxml/commit/b40f123e4fd8b5534013b6eded48dc119915a129))
- **docx:** Operate content controls around table cells ([4bfb1a8](https://github.com/ChristopherVR/ooxml/commit/4bfb1a800934130e984763aaae32396fd23dfea3))
- **docx:** Fill legacy form fields and honour forms protection ([8ab5a99](https://github.com/ChristopherVR/ooxml/commit/8ab5a99954432d51ddf8df92c4d6a7dba619de8e))
- **docx:** Leave sections excluded from forms protection editable ([28600f5](https://github.com/ChristopherVR/ooxml/commit/28600f5421c2e164fe079907e5aec05e469773c3))
- **docx:** Show the STYLEREF text of each page in Print Layout headers ([b82021a](https://github.com/ChristopherVR/ooxml/commit/b82021a062724fcdf96a45e88423f21b9382f380))
- **docx:** Refresh header STYLEREF results when a document is saved ([31837c3](https://github.com/ChristopherVR/ooxml/commit/31837c3fb1e2a7cb71e4ac7329381fbe9e486ef1))
- **docx:** Update body STYLEREF fields with Update Fields ([985b721](https://github.com/ChristopherVR/ooxml/commit/985b7216deea2cfc274a9294d578c696d39733ee))
- **docx:** Read the STYLEREF switch Word writes for chapter numbers ([2063b4f](https://github.com/ChristopherVR/ooxml/commit/2063b4f1415b532ee2e533ebfcc418a5ac603427))
- **xlsx:** Compare two workbooks by cells and structure ([cedde1b](https://github.com/ChristopherVR/ooxml/commit/cedde1bb090aacac7de7721e4f50031d2d8e617b))
- **xlsx:** Add a plain-text report of a workbook comparison ([b132bf4](https://github.com/ChristopherVR/ooxml/commit/b132bf4c3dd4e5cbba41531091ebc44e1f4b6902))
- **docx:** List the fill-in fields of a form with ids and labels ([9f672db](https://github.com/ChristopherVR/ooxml/commit/9f672dbf19e05cd9a4215a5212e8451d3a69aea2))
- **docx:** Fill form fields by id ([ea285e9](https://github.com/ChristopherVR/ooxml/commit/ea285e94d4d7b3fded55c9bc9a3a0d09b065c17a))
- **automation:** List and fill Word form fields from document bytes ([548ca2d](https://github.com/ChristopherVR/ooxml/commit/548ca2d53b2828bc69681e154e8bd668cfacf90b))

### Bug Fixes

- **chart:** Scale and draw series with very many points ([778a3e9](https://github.com/ChristopherVR/ooxml/commit/778a3e973b5d948e15ef2d0de6ac70a4243feb5a))
- **xlsx:** Save, group and navigate sheets with very many rows ([559d7d6](https://github.com/ChristopherVR/ooxml/commit/559d7d6cf6d1f00990f36343a565b1ab8ba980aa))
- **xlsx:** Insert timelines over pivots with many date records ([f79fb65](https://github.com/ChristopherVR/ooxml/commit/f79fb65308b1fbe23c634f19b400270e01d65ac9))
- **xlsx:** Write blocks taller than the call argument limit ([dc05f79](https://github.com/ChristopherVR/ooxml/commit/dc05f79779cb0555f2a42cf13e36f676bcfa1a7c))
- **xlsx:** Clear stored spills that their anchor no longer covers ([6e84afc](https://github.com/ChristopherVR/ooxml/commit/6e84afc62a414b49d9df1af09f13adb055b8ac92))
- **docx:** Add no folder entries to a saved package ([9688c30](https://github.com/ChristopherVR/ooxml/commit/9688c30da93a3a73436a248b59b5a64d71fc7585))
- **docx:** Keep the order of header and footer references on save ([5af3699](https://github.com/ChristopherVR/ooxml/commit/5af3699326b80300b537f19f839f2def2932b3cf))
- **docx:** Write only the section properties that changed ([01545a8](https://github.com/ChristopherVR/ooxml/commit/01545a89787cafe9602ba9fb453a181a18e9ca9e))
- **docx:** Declare no namespace again in rebuilt properties ([4a3a4da](https://github.com/ChristopherVR/ooxml/commit/4a3a4da7727d5bf7b2e417eead964abca49e8ca4))
- **xml:** Keep the line break after the declaration of a part ([9af3e23](https://github.com/ChristopherVR/ooxml/commit/9af3e23b3c33368027123f1363c9d1236cc8230e))
- **docx:** Save edited row and cell content controls ([735edbc](https://github.com/ChristopherVR/ooxml/commit/735edbc689a7a6581b22c1527ee04f0f5ca9cf02))
- **docx:** Keep a rebuilt content control's end properties ([dbff2b2](https://github.com/ChristopherVR/ooxml/commit/dbff2b27717f3783c3def6f29ee9b6a8276886b2))
- **docx:** Open a document whose glossary part cannot be read ([bd0aaa5](https://github.com/ChristopherVR/ooxml/commit/bd0aaa564dabfb88570bce7956c9f5da4fa06106))
- **docx:** Save the block controls around a table cell's paragraphs ([a407dd4](https://github.com/ChristopherVR/ooxml/commit/a407dd4fa4ea566ab6a91974b7c62114d5781da5))
- **docx:** Remove a temporary control when a picture in it is replaced ([cb7850e](https://github.com/ChristopherVR/ooxml/commit/cb7850e8644e34fee759aa19ee10de23f6def761))
- **docx:** Keep the caret in a cell whose control content is replaced ([e23bf7a](https://github.com/ChristopherVR/ooxml/commit/e23bf7aabcdd1dd192105df2ae7693a149d8fa52))
- **docx:** Remove a temporary control only when its content changes ([27eeae7](https://github.com/ChristopherVR/ooxml/commit/27eeae75903c975593f9f1f223cc7bd01be2b3f2))
- **docx:** Keep locked controls in a cell whose content is replaced ([45eee25](https://github.com/ChristopherVR/ooxml/commit/45eee25de9efa9b78d4507cf4a9d480e33f68e53))
- **docx:** Delete a selection that runs past a control in full ([ed3a6b7](https://github.com/ChristopherVR/ooxml/commit/ed3a6b7c881fc266b1321d7524d73b7a5686e987))
- **docx:** Type over an inline placeholder at its edge in a control ([4cc63eb](https://github.com/ChristopherVR/ooxml/commit/4cc63eb0dc025e1d0529d672943117331f60938c))
- **docx:** Give a pasted content control a new id ([c55cbe9](https://github.com/ChristopherVR/ooxml/commit/c55cbe996102c380e38ad2bb0ff1d77655f3d8a9))
- **docx:** Unwrap a control the model takes off all its rows or cells ([044a57c](https://github.com/ChristopherVR/ooxml/commit/044a57c4c37c8b3be4bd852ce955ef405631f9b2))
- **docx:** Keep the range markers beside a control when unwrapping it ([766c2ce](https://github.com/ChristopherVR/ooxml/commit/766c2cee1493198309c23260649483f59869d323))
- **docx:** Remove a temporary row control once its row is edited ([424008a](https://github.com/ChristopherVR/ooxml/commit/424008a96e539aafc8a52303ac2cfc2d5d0e01d0))
- **docx:** Tell apart temporary controls that share an id ([5683c6c](https://github.com/ChristopherVR/ooxml/commit/5683c6ce220a61c4be0ecfea9cf3b243a0dc50ea))
- **docx:** Renumber only pasted copies of controls, rows included ([17ca20f](https://github.com/ChristopherVR/ooxml/commit/17ca20f939b7c713385ad71a7475135e65d911ec))
- **docx:** Keep new content control ids within 32-bit integers ([6872925](https://github.com/ChristopherVR/ooxml/commit/687292504309523d7fe42b894280d855724636d2))
- **docx:** Remove a nested block control with a cell's replaced content ([f7b4640](https://github.com/ChristopherVR/ooxml/commit/f7b4640545107f987618676a3fee7241c4defd4f))
- **docx:** Remove a temporary control edited between its parts ([9d42d7f](https://github.com/ChristopherVR/ooxml/commit/9d42d7fe003bac124b46c49933a6ece37ced1e62))
- **docx:** Type over a cell placeholder only within its cell ([983a421](https://github.com/ChristopherVR/ooxml/commit/983a4212394db5b1b7c951ce38016b922d1c77f2))
- **docx:** Keep a deleted temporary control out of the text typed next ([7e2c1dc](https://github.com/ChristopherVR/ooxml/commit/7e2c1dc12850fc7fdefa954d17b63eca6b12d30d))
- **docx:** Write a new control's verbatim properties only as properties ([a5db4d4](https://github.com/ChristopherVR/ooxml/commit/a5db4d452204fe08f993e69e4fb4104bdd4aeea6))
- **docx:** Take a temporary control off only itself when ids repeat ([7cc7c34](https://github.com/ChristopherVR/ooxml/commit/7cc7c346bbfada55667285a52adf9c0f6159e259))
- **docx:** Keep a new control's id clear of the ids of row controls ([459fe27](https://github.com/ChristopherVR/ooxml/commit/459fe273f82b96cf9bea0ee2a4e07640fb69eff4))
- **docx:** Remove a temporary block control that holds a table ([5270cfa](https://github.com/ChristopherVR/ooxml/commit/5270cfa67fb8b0935588e92e22ba8b298c7549c3))
- **docx:** Keep stored formatting through a temporary control's removal ([39b20ea](https://github.com/ChristopherVR/ooxml/commit/39b20ea1236408485184005e5a4bd941d902cef7))
- **docx:** Give a row or cell another control in a w:sdt of its own ([18f11b6](https://github.com/ChristopherVR/ooxml/commit/18f11b6594f9d63b9bbc0a03e3d5f7769b774a77))
- **docx:** Write only valid content control properties ([9003efd](https://github.com/ChristopherVR/ooxml/commit/9003efd9d2391b91d5d52b9892d171d25e539fbc))
- **docx:** Take a content control from pasted HTML only when it is one ([72e8a53](https://github.com/ChristopherVR/ooxml/commit/72e8a5310895e2ea5022921ec768b7c718adc6fa))
- **docx:** Keep a control in one with locked content from changing ([a40013a](https://github.com/ChristopherVR/ooxml/commit/a40013a60f84c1829c3b6a35787aa7e97f41d25e))
- **docx:** Check pasted cell, row and block controls like inline ones ([5556931](https://github.com/ChristopherVR/ooxml/commit/5556931fcfc309ab3811e21e23468fa1b01d823d))
- **docx:** Let a node whose control may not be deleted take attributes ([2f70b3f](https://github.com/ChristopherVR/ooxml/commit/2f70b3f3858322eb4d4ea033fdc61908dd632280))
- **docx:** Apply the lock of a control around a table to the table ([953dbf5](https://github.com/ChristopherVR/ooxml/commit/953dbf5a4d191b443ee6ba756723503e69e252d9))
- **docx:** Keep locked row and table controls when replacing a cell ([6d5aeae](https://github.com/ChristopherVR/ooxml/commit/6d5aeaed5f9c88327476bf6c6635cebeb257a23d))
- **docx:** Leave range markers out of what a row or cell control holds ([6107aba](https://github.com/ChristopherVR/ooxml/commit/6107aba69ad28ef2d4444f506a9d139ae0e92224))
- **docx:** Join cells given one control whose id Word cannot read ([18b7511](https://github.com/ChristopherVR/ooxml/commit/18b7511490543d08e406cd9b7673d1567f6f07ad))
- **docx:** Keep a control whose picked entry shows no text ([1571fa0](https://github.com/ChristopherVR/ooxml/commit/1571fa0646868f5ad83bbe4498153c3c69e092c3))
- **docx:** Undo commands on a control that may not be deleted ([53cd8dd](https://github.com/ChristopherVR/ooxml/commit/53cd8dd924cdc2cd9af6306c581dbdcff079d3fd))
- **docx:** Apply a cell control's rules to a paste ([4a59e85](https://github.com/ChristopherVR/ooxml/commit/4a59e85b592c21601e57b9ee28e537fd1b108245))
- **docx:** Keep the properties of a control written anew as another kind ([98d121d](https://github.com/ChristopherVR/ooxml/commit/98d121d17b0b4649cced59f795a7f55bae538151))
- **docx:** Take and write only values of a control's own kinds ([cd0d434](https://github.com/ChristopherVR/ooxml/commit/cd0d434b0c03f4b691dd412d2ac9fdaf6faddbe1))
- **docx:** Let undo and redo take back a pasted locked control ([d768209](https://github.com/ChristopherVR/ooxml/commit/d76820901174729d6ec4f6422f42c250338532a4))
- **docx:** Take no pasted control nested deeper than a control is ([0578920](https://github.com/ChristopherVR/ooxml/commit/057892030a4a39478d50cba97fa1f528bb5857d9))
- **docx:** Edit paragraphs that hold proofing marks, symbols and smart tags ([8d2b489](https://github.com/ChristopherVR/ooxml/commit/8d2b48976c96322b0be46a5ee1a716d8f500fcb3))
- **docx:** Read smart tags and custom XML inside simple fields ([91a8d5b](https://github.com/ChristopherVR/ooxml/commit/91a8d5b9c93ee68442ba24b5506dfa8add0aa28a))
- **docx:** Keep markers and smart tags in place in a rebuilt paragraph ([8ddce70](https://github.com/ChristopherVR/ooxml/commit/8ddce70d17ba25ff6a1887f91c907ecba70a9b1f))
- **docx:** Draw symbols that share their run with text in their font ([c824a19](https://github.com/ChristopherVR/ooxml/commit/c824a1962c53d692c2e72c818b911e88c56da561))
- **docx:** Read a w:sym code that is half a surrogate pair as no symbol ([f76c279](https://github.com/ChristopherVR/ooxml/commit/f76c27920bf1307a013805d47380e310ce91ff2c))
- **docx:** Lay out positional tabs at their alignment point ([00c188d](https://github.com/ChristopherVR/ooxml/commit/00c188d671ecf61bb65fee3fd4f778ac1387d213))
- **docx:** Honour content control locks around legacy form fields ([1a7e72f](https://github.com/ChristopherVR/ooxml/commit/1a7e72fd20b89022fc197ecce7cd4eab07c99ba2))
- **docx:** Fill form fields whose data is indented in the begin marker ([17081e8](https://github.com/ChristopherVR/ooxml/commit/17081e88f81f35becbeb09346d2bf4340306eb28))
- **docx:** Place section breaks after block content controls ([7669156](https://github.com/ChristopherVR/ooxml/commit/76691569dda561c03fdb2eb62b6eb51b32054e71))
- **docx:** Keep headers, footers and notes read-only under forms protection ([4be8e69](https://github.com/ChristopherVR/ooxml/commit/4be8e69364180135598db377a0e781fe7182909c))
- **docx:** Let attributes change in the editable parts of a protected form ([8758cc9](https://github.com/ChristopherVR/ooxml/commit/8758cc9c8655bc8ea0c97c32e01b603d3650f2a5))
- **docx:** Place unprotected sections of a protected form as layout does ([23e746c](https://github.com/ChristopherVR/ooxml/commit/23e746c857a46baf3ac2a11f9a610900c2a9212a))
- **docx:** Let the editor renumber split paragraphs in a protected form ([3e338e5](https://github.com/ChristopherVR/ooxml/commit/3e338e5017a620acad9b0b011e09c5ec91f5dcd3))
- **docx:** Keep disabled text fields read-only inside content controls ([91f250e](https://github.com/ChristopherVR/ooxml/commit/91f250e0d33c65afe45aa9567c021f6fba44e9bb))
- **docx:** Operate form check boxes and drop-downs from the keyboard ([b0eb45d](https://github.com/ChristopherVR/ooxml/commit/b0eb45dd8f780b0dd82b3d11fc2e5a309c7d6ff9))
- **docx:** Keep markers inside the tracked change they were read in ([ca4f3c9](https://github.com/ChristopherVR/ooxml/commit/ca4f3c98e371862efa5bd76ccf278d41d90058a6))
- **text:** Keep unchanged words between far-apart edits in a diff ([3cf0b04](https://github.com/ChristopherVR/ooxml/commit/3cf0b0495a5c3c2d339b22c5db2b547c14921fbd))
- **docx:** Align positional tabs to the closer of indent and margin ([3b8e024](https://github.com/ChristopherVR/ooxml/commit/3b8e024798d923607ea5bef4e6e644c77589ec2e))
- **docx:** Treat neighbouring unprotected sections as one range ([85793e5](https://github.com/ChristopherVR/ooxml/commit/85793e5b39a7369c5e24ef03e34714e3876f8723))
- **docx:** Keep section ends in place when the editor renumbers blocks ([bc1e71d](https://github.com/ChristopherVR/ooxml/commit/bc1e71da10da3c89b3ca9dd9225ff256fa204393))
- **docx:** Keep an empty range in order inside a tracked change ([1d2535a](https://github.com/ChristopherVR/ooxml/commit/1d2535a03dc3e3245b122591ce8fb74a6b388a42))
- **docx:** Keep a disabled field's formatting in a protected form ([c14cad8](https://github.com/ChristopherVR/ooxml/commit/c14cad8770fceb74ed3bfb3a181fb46169c2e655))
- **docx:** Keep markers at their characters in merged tracked runs ([ddab61b](https://github.com/ChristopherVR/ooxml/commit/ddab61b8e7da6bddb6e5048fd4fccdbe0a986dfd))
- **docx:** Keep a hyperlink's markup when its whole label is replaced ([57e059d](https://github.com/ChristopherVR/ooxml/commit/57e059d7b3fb4b3515277c81c0a83b4388dc4bc8))
- **docx:** Replace a blank text field's en spaces with typed text ([4e15c92](https://github.com/ChristopherVR/ooxml/commit/4e15c92e18d738550bf23f1db0060098553e1493))
- **docx:** Change run boundaries over identical unmodeled properties ([38c2ad8](https://github.com/ChristopherVR/ooxml/commit/38c2ad8e45f833c800f430c9b6ac5e58cd25ab1a))
- **docx:** Keep symbols apart from text with the same character ([dda1b99](https://github.com/ChristopherVR/ooxml/commit/dda1b99c8d4adefabd97e12cd1ba737e88be596b))
- **docx:** Keep sections in place when a field splits a paragraph ([c8ef570](https://github.com/ChristopherVR/ooxml/commit/c8ef5702f6504ee134721c11d50e7c67a6aef856))
- **docx:** Enforce document protection that leaves enforcement unsaid ([b5c6adb](https://github.com/ChristopherVR/ooxml/commit/b5c6adb018c4b3913d4cb850f181c2e4e32bf131))
- **docx:** Undo an edit to a text field past its maximum length ([9eee5fb](https://github.com/ChristopherVR/ooxml/commit/9eee5fb720ec5618823e22857cb95b092b618667))
- **docx:** Keep a form field's bookmark around it when the field moves ([42dcb86](https://github.com/ChristopherVR/ooxml/commit/42dcb86ba0f9a2668c22f8463cc02ab4272df847))
- **docx:** Keep a pasted tab ordinary beside a positional tab ([e31ea3a](https://github.com/ChristopherVR/ooxml/commit/e31ea3a1f8fef35058c6dd239bac8a88f0555cb7))
- **docx:** Keep a protected form's fields from taking formatting ([e137fb6](https://github.com/ChristopherVR/ooxml/commit/e137fb69fbfee30f6e7661d34208deeb85078cfb))
- **docx:** Follow a form field's bookmark that opens after its begin ([47f8b2b](https://github.com/ChristopherVR/ooxml/commit/47f8b2b5ab7435d590e07adcb2786feab5c928bd))
- **docx:** Count an empty paragraph in the style as a STYLEREF match ([dd05b86](https://github.com/ChristopherVR/ooxml/commit/dd05b86a22c395c71bb47b8a604b285585b2d261))
- **docx:** Read tables nested in a cell in reading order for STYLEREF ([1af2735](https://github.com/ChristopherVR/ooxml/commit/1af27351529c85a62331370fd8854c9a40900dca))
- **docx:** Keep a STYLEREF \t result when words stand between levels ([f26b216](https://github.com/ChristopherVR/ooxml/commit/f26b216a42906aa7d3e7de533d381af019e34bb2))
- **docx:** Update STYLEREF fields in tables nested in a cell ([651fd55](https://github.com/ChristopherVR/ooxml/commit/651fd55087642abef402035826061ad56e94a81b))
- **docx:** Keep the typed error of an invalid model with STYLEREF fields ([3480b3a](https://github.com/ChristopherVR/ooxml/commit/3480b3aa3b8d37f9ea0e4dcf18c8e7062124468b))
- **docx:** Number an unnumbered STYLEREF match 0, as Word does ([a0eb495](https://github.com/ChristopherVR/ooxml/commit/a0eb495adecbdb740a0273dfb176bb736a7c09b4))
- **docx:** Format a CHARFORMAT STYLEREF result like its field name ([8030a2f](https://github.com/ChristopherVR/ooxml/commit/8030a2f04c5e10aa1c9d37ffae987d1796d676fb))
- **docx:** Count a paragraph box without lines as shown on its page ([3b91863](https://github.com/ChristopherVR/ooxml/commit/3b918630519025cbdf75a136fb672e94eb1bdfc2))
- **docx:** Find STYLEREF text around a page by the pages that show it ([3fbe54c](https://github.com/ChristopherVR/ooxml/commit/3fbe54cbf9adcdd5e25573db54d45e3315c8be39))
- **docx:** Clear a body STYLEREF result in Update Fields ([b173820](https://github.com/ChristopherVR/ooxml/commit/b1738206c0606a95963611fad547d8eb466560c8))
- **docx:** Reformat a CHARFORMAT STYLEREF result whose text is current ([c5687d5](https://github.com/ChristopherVR/ooxml/commit/c5687d5bcc7afdae91d3f1f721e2b22cfa08ef98))
- **docx:** Number list paragraphs in tables nested in a cell ([c96f229](https://github.com/ChristopherVR/ooxml/commit/c96f2293d12eafd1b11e23e48eebd6b4d7330959))
- **docx:** Refresh STYLEREF results on save part by part ([1043f94](https://github.com/ChristopherVR/ooxml/commit/1043f9408b9cf0c775aef311463cce6047e74133))
- **docx:** Keep field results out of text a revision deletes ([c6b59e4](https://github.com/ChristopherVR/ooxml/commit/c6b59e4d16aa87f65e1ea7b5716d162e01e5143d))
- **docx:** Treat text hidden through its style as hidden in STYLEREF ([af198c2](https://github.com/ChristopherVR/ooxml/commit/af198c21d842676b68ceacbd7352a1394a9a0c27))
- **docx:** Read a STYLEREF style that comes after its switches ([ba5191a](https://github.com/ChristopherVR/ooxml/commit/ba5191a1d761c46b58fb7d186f36992883a67377))
- **docx:** Leave a paragraph a tracked change deletes out of STYLEREF ([c1bcfe6](https://github.com/ChristopherVR/ooxml/commit/c1bcfe6125ba351ad4beb1defcb06e626037be0b))
- **docx:** Give a STYLEREF \n in a numbered paragraph its own number ([3dae245](https://github.com/ChristopherVR/ooxml/commit/3dae24520f018db7ab664978e08e794254868141))
- **pptx:** Render weight-named font families at their weight ([7188288](https://github.com/ChristopherVR/ooxml/commit/718828829585145b461742c04e030fe3fbbd2d52))
- **color:** Require each palette variant to name its step ([70d7d9a](https://github.com/ChristopherVR/ooxml/commit/70d7d9a949e9dc5289bacf33bf63bd9af4c9f258))
- **docx:** Import verbatim control properties through importElement ([6f02c7e](https://github.com/ChristopherVR/ooxml/commit/6f02c7e8d4be2d7787ff1566de336f9fe1b95a5f))
- **color:** Require each palette variant to name its step ([3135862](https://github.com/ChristopherVR/ooxml/commit/3135862978fd352498d977335777555ba519428a))
- **docx:** Drop a label's trailing colon without a pattern ([8860b15](https://github.com/ChristopherVR/ooxml/commit/8860b15ee32640dc75305dc8cd175fb83c3fe883))
- **docx:** Leave text hidden through its style out of form labels ([2b9c5a2](https://github.com/ChristopherVR/ooxml/commit/2b9c5a2ed1dd8db6bb1a2ee35585c2f42d86a850))
- **docx:** Write a new content control's properties only when valid ([3a8e23a](https://github.com/ChristopherVR/ooxml/commit/3a8e23ac2df41db82357568749240592eab361a5))
- **docx:** Keep foreign verbatim properties out of pasted controls ([9ee6664](https://github.com/ChristopherVR/ooxml/commit/9ee666419ca8fdcf4bea3a3f58753011057a798f))
- **docx:** Keep a copied control's binding and check extension types ([19e6e78](https://github.com/ChristopherVR/ooxml/commit/19e6e785a34c94e0e8bf86e92dc4def22e0972fc))
- **docx:** Take and write only values of a control's own kinds ([cb4b8e1](https://github.com/ChristopherVR/ooxml/commit/cb4b8e148a3666da6abf9175cc3760699a38a4a0))
- **pptx:** Render flowChart connector presets as shapes, not lines ([08399ae](https://github.com/ChristopherVR/ooxml/commit/08399aea49db3270d6ef9c9c163c0ea3b6f8e23f))
- **pptx:** Paint picture fills inherited from layout placeholders ([70bf7f9](https://github.com/ChristopherVR/ooxml/commit/70bf7f91beda9ae8e4d70472134222afc873dc68))

### Performance

- **xlsx:** Copy cells cheaply in undo records ([434c330](https://github.com/ChristopherVR/ooxml/commit/434c330152e9be61cd575b63754b34229c223f6d))
- **xlsx:** List and recalculate a bulk write's cells per block ([e14aab9](https://github.com/ChristopherVR/ooxml/commit/e14aab97cdb409669194064eb7abb77cc312c4fc))
- **xlsx:** Copy each format once when copying a block ([cbd9a10](https://github.com/ChristopherVR/ooxml/commit/cbd9a104732e0140d10c8b0dc18c408d3dd602cd))
- **xlsx:** Look up changed cells only for formulas that may spill ([611e237](https://github.com/ChristopherVR/ooxml/commit/611e2376e4d074c03e01a9a2730c84d0282ac6c2))
- **xlsx:** Clear a sparse stored spill by its stored cells ([a1f04d5](https://github.com/ChristopherVR/ooxml/commit/a1f04d588169a00417db55091a6efc40e5c0c45b))
- **docx:** Rescan only the edited text for symbol fonts ([9041c6b](https://github.com/ChristopherVR/ooxml/commit/9041c6b72564d43ebd11e631995f7c877b36757d))

### Refactor

- **docx:** Move the section property writer into its own module ([7450804](https://github.com/ChristopherVR/ooxml/commit/74508048f93e37beba499868a1fbdc00bbf11c24))

### Documentation

- **xlsx:** Describe csv typing as an approximation of excel ([8b38142](https://github.com/ChristopherVR/ooxml/commit/8b381425d33ee98a11ca8b1e8ea9d5ba91845503))
- **xlsx:** Document workbook comparison ([92db6c7](https://github.com/ChristopherVR/ooxml/commit/92db6c7439988394da7ad1aa1128a06e92c7e1ad))

### Testing

- **xlsx:** Pin the cached range end for a date-time maximum ([7e126ce](https://github.com/ChristopherVR/ooxml/commit/7e126ce9c177a4e6fffccaab818f7f392a7397fa))
- **xlsx:** Make the element limit test reach the full parser again ([8a42ee8](https://github.com/ChristopherVR/ooxml/commit/8a42ee8b47ab45aabc983e0de35ef7aac41fe40b))
- **docx:** Pin STYLEREF heading levels to the built-in heading names ([e42cd63](https://github.com/ChristopherVR/ooxml/commit/e42cd63930beaf3c032ae5c978512e0af765b448))

## [1.20.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.20.0) - 2026-10-21

### Features

- **xlsx:** Expose composable spreadsheet building blocks ([9aef499](https://github.com/ChristopherVR/ooxml/commit/9aef499f76196be6dd008ae873ac8f0d4caeeb58))
- **visio:** Expose canvas building blocks and host permissions ([e464a1c](https://github.com/ChristopherVR/ooxml/commit/e464a1c0399434430925e0e59b430b5788f5d8d9))
- **xlsx:** Expand workbook inspection and cell reads ([#117](https://github.com/ChristopherVR/ooxml/issues/117)) ([481bcc2](https://github.com/ChristopherVR/ooxml/commit/481bcc265611822a628c121f00897ec6b3e95cd6))
- **pptx:** Add layout-inheriting text placeholders ([#113](https://github.com/ChristopherVR/ooxml/issues/113)) ([a9bb079](https://github.com/ChristopherVR/ooxml/commit/a9bb07957fcf9aa1a12406911726211b415c690a))

### Bug Fixes

- **ui:** Harden composition lifecycle and consumer checks ([6aa52be](https://github.com/ChristopherVR/ooxml/commit/6aa52be88175ca4f46f6f3b1c223ef06a8f1f915))

### Performance

- **xlsx:** Record only changed row and column dimensions ([#115](https://github.com/ChristopherVR/ooxml/issues/115)) ([143cfd5](https://github.com/ChristopherVR/ooxml/commit/143cfd51c4a63aec19aae6df41993b6127378663))
- **xlsx:** Keep comments and cdata on the XML fast path ([#119](https://github.com/ChristopherVR/ooxml/issues/119)) ([ee7212a](https://github.com/ChristopherVR/ooxml/commit/ee7212afc204e15b5977d215f5cc61c99aaf4bc3))

### Dependencies

- **deps:** Update dev tools and fix svelte type checks ([#107](https://github.com/ChristopherVR/ooxml/issues/107)) ([550eac2](https://github.com/ChristopherVR/ooxml/commit/550eac24d469d14245a0ac71789e01c8d5fb50f9))

### Chores

- Integrate latest main before full verification ([3bec814](https://github.com/ChristopherVR/ooxml/commit/3bec814857ae675f31eba14fcec59a3db4e1cb77))

## [1.19.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.19.2) - 2026-10-20

### Bug Fixes

- **docx:** Preserve text edits in irregular table rows ([#111](https://github.com/ChristopherVR/ooxml/issues/111)) ([fd234d7](https://github.com/ChristopherVR/ooxml/commit/fd234d7484f52a0a564a7e5e71306b5f811064b2))

### Performance

- **xlsx:** Avoid repeated grid scans when fitting rows ([#112](https://github.com/ChristopherVR/ooxml/issues/112)) ([2dc34f7](https://github.com/ChristopherVR/ooxml/commit/2dc34f77d65deedb5ad13efe5f051ea7f1805808))

### Dependencies

- **deps:** Update emf-converter to 4.21.0 ([#106](https://github.com/ChristopherVR/ooxml/issues/106)) ([641ae31](https://github.com/ChristopherVR/ooxml/commit/641ae316b7de07ed6f3c5158e8b154c991e1b719))

## [1.19.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.19.1) - 2026-10-20

### Bug Fixes

- **deps:** Adapt error codes and iframe tests to updated dependencies ([550070e](https://github.com/ChristopherVR/ooxml/commit/550070e59951c339a67c7e3d6920b9b7ac0c1004))
- **pptx:** Link authors when saving first-time comments ([6b6ea0d](https://github.com/ChristopherVR/ooxml/commit/6b6ea0d149f8073ddf10d44454795e54386e5ec3))

### Dependencies

- **deps:** Merge dependency update ([#92](https://github.com/ChristopherVR/ooxml/issues/92)) ([ce1b012](https://github.com/ChristopherVR/ooxml/commit/ce1b01215261b162443f903e1a2086ed4110ae28))
- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))
- **deps:** Merge dependency update ([#95](https://github.com/ChristopherVR/ooxml/issues/95)) ([0926eac](https://github.com/ChristopherVR/ooxml/commit/0926eac231cb49749433fb7a9eac870cc7996e52))
- **deps:** Merge dependency update ([#96](https://github.com/ChristopherVR/ooxml/issues/96)) ([33b0ebb](https://github.com/ChristopherVR/ooxml/commit/33b0ebb6a16b4448d577c169c79187466d4148a3))

### Chores

- **deps-dev:** Bump @types/node ([df60cd7](https://github.com/ChristopherVR/ooxml/commit/df60cd7aff4efb7006ac0a6cfd1c69b08230a011))
- **deps-dev:** Bump tsup from 8.5.1 to 9.0.0 ([9835e80](https://github.com/ChristopherVR/ooxml/commit/9835e8038a9a271e37e286e524496513910553a6))
- **deps-dev:** Bump the development-minor-and-patch group with 14 updates ([1cdc60b](https://github.com/ChristopherVR/ooxml/commit/1cdc60b0f0a591304a2240ffcb550eccfd8d2487))
- **pptx:** Integrate current main with dependency updates ([713239b](https://github.com/ChristopherVR/ooxml/commit/713239b3b9166d0010dcd058ea758c11ff2425c0))

## [1.19.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.19.0) - 2026-10-17

### Features

- **chart:** Honor native markers throughout shared rendering ([a7aeee1](https://github.com/ChristopherVR/ooxml/commit/a7aeee171248fc565b7bc9d611e054b659954243))

### Bug Fixes

- **chart:** Apply shared axis scaling to radar charts ([4acbf15](https://github.com/ChristopherVR/ooxml/commit/4acbf1556d8ac89371f736a8b7b5f7a5aedc3609))

### Refactor

- **chart:** Share axis ranges and native gridline rendering ([8c26687](https://github.com/ChristopherVR/ooxml/commit/8c26687a7b3196955df602198528adc83fce259b))
- **chart:** Share pptx marker geometry with neutral rendering ([b246d63](https://github.com/ChristopherVR/ooxml/commit/b246d6315c32412298508d9b2a944b0c9bafc26b))

## [1.18.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.18.0) - 2026-10-17

### Features

- **xlsx:** Write and read a chart workbook for embedding ([393bb58](https://github.com/ChristopherVR/ooxml/commit/393bb58720f7f268f43fb5c42331b8cd9a46f893))
- **visio:** Embed an Excel chart as an object, as Insert > Chart does ([e1f7db3](https://github.com/ChristopherVR/ooxml/commit/e1f7db3de67b3afe5aa76d6c980c5ce59e3476b1))
- **visio:** Place a new chart object on the page ([c3a8a01](https://github.com/ChristopherVR/ooxml/commit/c3a8a0115b2db7366c0ff570bd0b2a67ae89f8a4))
- **chart:** Draw data labels in the shared chart renderer ([32142ef](https://github.com/ChristopherVR/ooxml/commit/32142ef32c392548fdf5f79aa9d7f659356fac70))

### Bug Fixes

- **chart:** Preserve data label inheritance and source formats ([e92e5eb](https://github.com/ChristopherVR/ooxml/commit/e92e5eb80b098f6331fa21a5dc284018d56ec686))

### Documentation

- **chart:** Gap report and harness for the shared chart renderer ([b166877](https://github.com/ChristopherVR/ooxml/commit/b1668770e66e241704dd5ec9a04d2813e522acc4))

## [1.17.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.17.0) - 2026-10-17

### Features

- **ui:** Honour undo levels and ScreenTip style in visio options ([1ce62d3](https://github.com/ChristopherVR/ooxml/commit/1ce62d39fe7c7101f91ce93d4a34e612830946be))
- **xlsx:** Draw a chart from typed grid data ([ff985b6](https://github.com/ChristopherVR/ooxml/commit/ff985b69c0018da39896e21504132d07bf8fdf5b))

### Bug Fixes

- **xml:** Close the remaining code scanning alerts ([6d85151](https://github.com/ChristopherVR/ooxml/commit/6d85151b85dc65fe03c26079c8e76e162dffd5c5))
- **xlsx:** Look sheets up with at so no key reaches a prototype ([41a0b04](https://github.com/ChristopherVR/ooxml/commit/41a0b04fd2dde8a6f5ece20ac46b595c3a130108))

## [1.16.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.16.1) - 2026-10-16

### Bug Fixes

- **xml:** Close code scanning alerts and steady the docx presence spec ([442b249](https://github.com/ChristopherVR/ooxml/commit/442b249cce1de4fdd228d073b394d741733f4c08))

## [1.16.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.16.0) - 2026-10-16

### Features

- **xlsx:** Values on rows under a PivotTable report filter ([cd095b0](https://github.com/ChristopherVR/ooxml/commit/cd095b0dd28a46e9e16c8f8a3bf46d47c7a8b9b8))
- **xlsx:** Reproduce PivotTable captions in the language of the file ([be4c874](https://github.com/ChristopherVR/ooxml/commit/be4c874a79b30be170a69ef0b68efa356b74a79a))
- **pptx:** Apply theme effect schemes and background styles ([998bb4b](https://github.com/ChristopherVR/ooxml/commit/998bb4b52c0d3606d0d4c26a0f26c970b4684b69))
- **pptx:** Resolve a style against another effect scheme ([6db6a07](https://github.com/ChristopherVR/ooxml/commit/6db6a0777b8a82a787f8a78015b34513df4c2f0a))
- **pptx:** Expose a package's real theme variants on the handler ([f176c1b](https://github.com/ChristopherVR/ooxml/commit/f176c1bcd2974005606fcc3588a5d3620de414b8))
- **pptx:** Save chart picture, pattern, gradient, 3-D and text formats ([979e444](https://github.com/ChristopherVR/ooxml/commit/979e444af1cab982a8e1ea2360325646efeec6d4))

### Bug Fixes

- **pptx:** Write ChartEx bin size and count as val attributes ([457df13](https://github.com/ChristopherVR/ooxml/commit/457df13ae1c4bafbb61d53b41a3078e7a74556b4))
- **pptx:** Write animation graphicEl children in the DrawingML namespace ([9b7604d](https://github.com/ChristopherVR/ooxml/commit/9b7604d417495e0c89342207ffb2ed4afc5906ab))
- **pptx:** Mirror chart text colour edits and keep automatic label angles ([c8f12c3](https://github.com/ChristopherVR/ooxml/commit/c8f12c3b17c0fc778fc5905a4d8ba8c0f2fd0e2a))
- **pptx:** Keep a chart title's size when its text format is edited ([67b433c](https://github.com/ChristopherVR/ooxml/commit/67b433c3dd77863e37201441c851dc9b1e9b60ec))
- **pptx:** See data labels of a second chart group of the same type ([06e7685](https://github.com/ChristopherVR/ooxml/commit/06e7685fd29fc9c5239c8161aa0963e09378adbc))
- **pptx:** Sync a slide background style over collaboration ([32bc400](https://github.com/ChristopherVR/ooxml/commit/32bc400deccd8a4c53df15aec8e4c1245b392654))
- **pptx:** Center resized slide content in pixel coordinates ([#86](https://github.com/ChristopherVR/ooxml/issues/86)) ([fbbdfee](https://github.com/ChristopherVR/ooxml/commit/fbbdfee9dfe9b718ed02f0847e2a1197661f2f33))
- **pptx:** Preserve picture source crops in svg exports ([#88](https://github.com/ChristopherVR/ooxml/issues/88)) ([97c254c](https://github.com/ChristopherVR/ooxml/commit/97c254c159e7486e77b425be95bbe17aca0955ea))

### Testing

- **pptx:** Make the generated e2e decks open in PowerPoint ([036e0eb](https://github.com/ChristopherVR/ooxml/commit/036e0eb438d8664230afe7fe2de44a1e9365fb9b))

## [1.15.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.15.0) - 2026-10-16

### Features

- **teams:** Probe, check and end sessions from the client ([dab7217](https://github.com/ChristopherVR/ooxml/commit/dab7217734937ba5851796ad89e3a9c25d640d37))
- **teams:** Repeat meetings monthly and yearly ([3eb2406](https://github.com/ChristopherVR/ooxml/commit/3eb2406733a4a19791af8b0196357288d0e7b04e))
- **xlsx:** Custom subtotals on PivotTable column fields ([0b86fd5](https://github.com/ChristopherVR/ooxml/commit/0b86fd53c1743829a003cad9ea1a22d1d9f1eb0e))
- **xlsx:** Values on rows with any number of PivotTable fields ([c227437](https://github.com/ChristopherVR/ooxml/commit/c2274377b777b4ca98bc95a1b7d2d5c6ab97f634))
- **xlsx:** Date groups with report filters, hidden items and timelines ([a072ec6](https://github.com/ChristopherVR/ooxml/commit/a072ec6c8b1e8983686abe15f064eee8ed0741e9))
- **xlsx:** Read date group labels written in another language ([eccf72b](https://github.com/ChristopherVR/ooxml/commit/eccf72bb7ce10e827b0b9a360dbdf60c7489cfb3))
- **geometry:** Add the PowerPoint shape gallery catalogue ([2fb79e2](https://github.com/ChristopherVR/ooxml/commit/2fb79e2717cf9094a2692d04503e95cc09453467))
- **pptx:** Add chart element format and options edit operations ([d99f83e](https://github.com/ChristopherVR/ooxml/commit/d99f83ec63e6251c4aaa39a31ad5d2f4d2659b54))
- **pptx:** Read chart element formats from each element's spPr on load ([46620d0](https://github.com/ChristopherVR/ooxml/commit/46620d08e61513234ba61964e4f751d679fcb28b))

### Testing

- **xlsx:** Accept column subtotals, row values and filtered date groups in Excel ([c8c0f05](https://github.com/ChristopherVR/ooxml/commit/c8c0f053fcaf956a2b2df2d52f2eac17558e53a0))

## [1.14.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.14.0) - 2026-10-15

### Features

- **ui:** Edit headers and footers in place on print layout pages ([9e985ae](https://github.com/ChristopherVR/ooxml/commit/9e985ae1a29ec3a23d9555d2b5e847bf143c3173))
- **chart:** Draw combination, waterfall and histogram charts ([ae79ec2](https://github.com/ChristopherVR/ooxml/commit/ae79ec24da6576a4497fcc4fce0b444be436dcc1))
- **xlsx:** Draw combination, waterfall and histogram charts ([753dea7](https://github.com/ChristopherVR/ooxml/commit/753dea7e9f64cb3641643df907d5faa35612cf11))
- **docx:** Draw combination, waterfall and histogram charts ([93ee797](https://github.com/ChristopherVR/ooxml/commit/93ee797ad2083797d13e28ef8bce6df9bbdc333e))
- **xlsx:** Add page layout and page break preview geometry ([ec99c5d](https://github.com/ChristopherVR/ooxml/commit/ec99c5d1078371b436f5b729bbcf71947f4295b0))
- **xlsx:** Add workbook view commands, icons and strings ([d956e46](https://github.com/ChristopherVR/ooxml/commit/d956e46817dc62d46f21bc471e87cbfda428e9dc))
- **xlsx:** Fill each page view page with the cells that fit ([c76fb60](https://github.com/ChristopherVR/ooxml/commit/c76fb606796a9734956427b5ae33120925583f67))
- **xlsx:** Add goal seek, data tables and scenarios to core ([a9ebe39](https://github.com/ChristopherVR/ooxml/commit/a9ebe39efc70d330b1ca34adfcc2b18568d9fff2))
- **xlsx:** Add what-if commands, strings and data table formula text ([532dedc](https://github.com/ChristopherVR/ooxml/commit/532dedc25aaa80df06ba75e77f90db0bd1d4acd6))
- **teams:** Add @mentions with a people picker in the composer ([2bb877a](https://github.com/ChristopherVR/ooxml/commit/2bb877a050dae027eb948e6d52b9b068934d927e))
- **teams:** Add an Activity feed derived from the shared chat data ([6f53e8a](https://github.com/ChristopherVR/ooxml/commit/6f53e8ac925ba74f7f4a4b1dec37c5b6849797ff))
- **teams:** Announce new activity with toasts, desktop notifications and sound ([fe29b00](https://github.com/ChristopherVR/ooxml/commit/fe29b006d4830489b19c097290b132a5581ddb98))
- **visio:** Read stencils and drop their masters into a drawing ([2920f00](https://github.com/ChristopherVR/ooxml/commit/2920f00549b4831486b85d6a91b2726de9daf83d))
- **ui:** Open stencil files and collect My Shapes in the visio pane ([fe6312f](https://github.com/ChristopherVR/ooxml/commit/fe6312fa2da5e0dab162acacfc28b3b91fb80bc2))
- **visio:** Reroute right-angle connectors around a shape moved onto them ([fb19bef](https://github.com/ChristopherVR/ooxml/commit/fb19befbba9c2483a6627f33c730ccbc6cfcc658))
- **visio:** Recalculate formulas that use rounding, power and log ([b2bc840](https://github.com/ChristopherVR/ooxml/commit/b2bc8408557c6b997fe8315c822929a334233b62))
- **teams:** Add private chats, meetings and a calendar to the client ([a6c2748](https://github.com/ChristopherVR/ooxml/commit/a6c2748426473192a2e16d37d3856f8accc27673))
- **chart:** Draw bubble, stock, treemap and sunburst charts ([94a4240](https://github.com/ChristopherVR/ooxml/commit/94a424016e799722b22205f87b17e68fd5d5f816))
- **xlsx:** Draw bubble, stock, treemap and sunburst charts from live cells ([11a4b36](https://github.com/ChristopherVR/ooxml/commit/11a4b367b7e116590c60d8eaccbf051563ae5c97))
- **docx:** Draw bubble, stock, treemap and sunburst charts ([7053196](https://github.com/ChristopherVR/ooxml/commit/7053196003aeb7ad83dcb546ffef0a179c4331c5))
- **chart:** Honour stock bar formatting, draw volume stock, pad bubble axes ([614fc42](https://github.com/ChristopherVR/ooxml/commit/614fc421668e707be211c37658ea76681174eed5))
- **ui:** Edit nested table cells on print layout pages ([8b2662b](https://github.com/ChristopherVR/ooxml/commit/8b2662b9dc87322311a02d7ec6e9b97164da6280))
- **docx:** Write moved floating pictures and report their frame origin ([e272dc4](https://github.com/ChristopherVR/ooxml/commit/e272dc42157b758b96bf4c425681fdb9597d9c89))
- **docx:** Paginate plain documents incrementally after an edit ([7521712](https://github.com/ChristopherVR/ooxml/commit/7521712eba07913ff90a9205d982cd7a434a95b9))
- **xlsx:** Read, filter and write table and pivot slicers ([9cb2c26](https://github.com/ChristopherVR/ooxml/commit/9cb2c26b7d8226707b0d299d313e145dd2934d5b))
- **xlsx:** Add slicer commands, helpers and strings ([bf7db7f](https://github.com/ChristopherVR/ooxml/commit/bf7db7f08d54279540d942230c15cec98f4f2193))
- **xlsx:** Read, filter and write PivotTable timelines ([e4282ec](https://github.com/ChristopherVR/ooxml/commit/e4282ec69e15ffbecdab8ce7371df2cffa801128))
- **ui:** Draw timelines and add Insert > Timeline with a Timeline tab ([c525b8c](https://github.com/ChristopherVR/ooxml/commit/c525b8c83558ba8c23ec0e2c772cae57bd4e26d8))
- **xlsx:** Group PivotTable fields by number ranges ([9a9014f](https://github.com/ChristopherVR/ooxml/commit/9a9014f0728979e09843d7781130e09b45cecfac))
- **docx:** Read, edit and write content controls ([42a22a3](https://github.com/ChristopherVR/ooxml/commit/42a22a3d6bdb0babf6eb82792173958d179cfeb4))
- **docx:** Edit and insert content controls in the editor ([b1f3fce](https://github.com/ChristopherVR/ooxml/commit/b1f3fce94392d32919e1bd716c1cc5c0b7bf8911))
- **math:** Draw equations as SVG and write them to Office Math ([cf5d2c0](https://github.com/ChristopherVR/ooxml/commit/cf5d2c09691aea6434a35a23d8116d888a79f320))
- **ui:** Insert and edit equations from LaTeX in the Word editor ([88b8964](https://github.com/ChristopherVR/ooxml/commit/88b8964bc1ff1509f37889f835dad75762f57268))
- **docx:** Read, draw and write drawing shapes ([7c39452](https://github.com/ChristopherVR/ooxml/commit/7c3945284384cf263c4a3c21f2b31bce57a43d82))
- **ui:** Draw and format shapes from Insert > Shapes ([b63e03c](https://github.com/ChristopherVR/ooxml/commit/b63e03c93b3dca7e9bb37946b9a841f7b779358c))
- **automation:** Build word, excel and powerpoint files from outlines ([eda7e9e](https://github.com/ChristopherVR/ooxml/commit/eda7e9e7e5739764d2de03f93a71dcb7a268a78a))
- **teams:** Share assistant chats over the collaboration engine ([439b0ce](https://github.com/ChristopherVR/ooxml/commit/439b0ce5010fb4722b955c3385eb087bf76301b1))
- **automation:** Suggest word edits as tracked changes, read all cells ([0fdfd3a](https://github.com/ChristopherVR/ooxml/commit/0fdfd3ae752d605db5171d66b2355f8a25e85054))
- **teams:** Add session sign-in and send the session on sockets ([1abb030](https://github.com/ChristopherVR/ooxml/commit/1abb03076b9cab5aa1332abed9701e7c30f8a519))
- **teams:** Recurrence end, exceptions and zoned ics export ([fa6a0b0](https://github.com/ChristopherVR/ooxml/commit/fa6a0b08c26dc838b81ce506b3eb1fb6f8f15871))
- **visio:** Evaluate string, angle, point and error ShapeSheet functions ([af348d0](https://github.com/ChristopherVR/ooxml/commit/af348d0f460e5372bbc16573f2827d5ddf9dfbb4))
- **visio:** Follow glued shapes with straight stencil line connectors ([0866862](https://github.com/ChristopherVR/ooxml/commit/0866862d13fef36c0140501b83a7aec605055e9a))
- **chart:** Model series trendlines and error bars ([bd885ce](https://github.com/ChristopherVR/ooxml/commit/bd885ce5b90b318232363f1d801461ef0bce053e))
- **chart:** Fit trendlines and size error bars in data space ([1023d6f](https://github.com/ChristopherVR/ooxml/commit/1023d6f892c4534af9e83b34c638c17d26079bef))
- **chart:** Draw trendlines and error bars in the shared painter ([dcce901](https://github.com/ChristopherVR/ooxml/commit/dcce901233ca1c72580b094c243125c837961b17))
- **xlsx:** Date groups, any number of row fields, custom subtotals ([5e86f18](https://github.com/ChristopherVR/ooxml/commit/5e86f18d1b8f635f8d4ead3b5e52b4a9ee91fc0e))
- **ui:** Group dates, nest row fields and set subtotals per field ([614ac2c](https://github.com/ChristopherVR/ooxml/commit/614ac2c5206fb8b5a7fe538cccbb31ef958d248d))
- **xlsx:** Nest any number of PivotTable column fields ([3b612e4](https://github.com/ChristopherVR/ooxml/commit/3b612e40ba69c5bbcbc2e2f80064f521f334d3e0))
- **ui:** Add further column fields in the PivotTable settings ([a8d540b](https://github.com/ChristopherVR/ooxml/commit/a8d540b4b5292349cca62ee21476d45352b1513e))

### Bug Fixes

- **opc:** Stop reporting digest checks as a valid signature ([9b3cac4](https://github.com/ChristopherVR/ooxml/commit/9b3cac4c24c3a2fd17693fa253ceea6970ca0fe7))
- **pptx:** Guard OCSP responder requests against SSRF ([b223ee0](https://github.com/ChristopherVR/ooxml/commit/b223ee0f2ac593f3b3ac2ec1c452724be4a3fd07))
- **opc:** Guard zip reads against forged sizes and unsafe part names ([39294e9](https://github.com/ChristopherVR/ooxml/commit/39294e97d434e7132a22e39930d262a0d0b281da))
- **digest:** Add timingSafeEqual and a 1M default spin limit ([5efd9c9](https://github.com/ChristopherVR/ooxml/commit/5efd9c95dcb2fa1b5e8d830333b9dce5d253e816))
- **crypto:** Harden agile and standard decryption against hostile input ([b1e7be5](https://github.com/ChristopherVR/ooxml/commit/b1e7be51db6aee35ed91dfcf143a3710d199308a))
- **xlsx:** Bound formula work from names, wildcards, regex and arrays ([bcc5ec4](https://github.com/ChristopherVR/ooxml/commit/bcc5ec43ee4357e9d7b7d571a9f8b83b896c4aa7))
- **xlsx:** Clamp pasted html table spans to the sheet and a cell budget ([776913c](https://github.com/ChristopherVR/ooxml/commit/776913ce48357c521cad63467d2c3ff6885c005c))
- **xlsx:** Mark spills over populated cells, not every row of the ref ([1de1d03](https://github.com/ChristopherVR/ooxml/commit/1de1d03e0e6663253231035dcf0a48d84a205726))
- **xlsx:** Scan carried chart and pivot tags in linear time ([365b079](https://github.com/ChristopherVR/ooxml/commit/365b079a614e29e2a59b813b68c111f31da4438d))
- **automation:** Confine file io by realpath with relative results ([3b0340f](https://github.com/ChristopherVR/ooxml/commit/3b0340fa753c4fa0c483c0715c19693bd5b40b4e))
- **xml:** Bound nesting depth and element count in parseXml ([7441f60](https://github.com/ChristopherVR/ooxml/commit/7441f60925b128da24500833aecb5936d24ee024))
- **opc:** Match attribute names exactly and accept single quotes ([9220247](https://github.com/ChristopherVR/ooxml/commit/922024743255a61b6f39b2ef231dffc533da8b14))
- **math:** Cap LaTeX and OMML nesting and use null-prototype sibling maps ([553cd72](https://github.com/ChristopherVR/ooxml/commit/553cd7232a58d7b979f5065e35d1e9f31c2d48fc))
- **docx:** Bound list start values and letter labels, use null-prototype maps ([4f9360a](https://github.com/ChristopherVR/ooxml/commit/4f9360ae98ea061a556db4a773bed3271dbc4052))
- **diagram:** Null-prototype layout variables and a presentation node budget ([572bf71](https://github.com/ChristopherVR/ooxml/commit/572bf71479c24389793db22dc80f056e090d4264))
- **teams:** Require wss on https pages and flag untrusted configs ([cd8cd90](https://github.com/ChristopherVR/ooxml/commit/cd8cd9058c70d5412c98c83fea14589f7d52cbe7))
- **xlsx:** Tidy float noise in the goal seek result ([3bbb9dd](https://github.com/ChristopherVR/ooxml/commit/3bbb9dd087f2e4ca611899e04e4a36a48b05e499))
- **teams:** Place notification toasts clear of the composer and show mentions without @ ([1e46d6d](https://github.com/ChristopherVR/ooxml/commit/1e46d6dc907b72cbe63359ae94cd4648b5e09697))
- **ui:** Polish the Chat and Calendar apps after a browser pass ([02e5543](https://github.com/ChristopherVR/ooxml/commit/02e5543393713280ab8c941094ef4a282c39ddfc))
- **chart:** Match Excel for extended chart text, labels and fills ([9514ba1](https://github.com/ChristopherVR/ooxml/commit/9514ba19f9624b205da5a64b2ea6f32ea3355d65))
- **xlsx:** Remove unused slicer imports ([3948cbc](https://github.com/ChristopherVR/ooxml/commit/3948cbc08a57b1cb4401a8e329935baf05078098))
- **xlsx:** Soften slicer and timeline style tints ([5640e20](https://github.com/ChristopherVR/ooxml/commit/5640e20a4345f47e2b6980858c80e37ff13d7f68))
- **xlsx:** Save sheets built without a view as the default view ([0246f49](https://github.com/ChristopherVR/ooxml/commit/0246f499ce835acea563c2b7ba120925cfe5b79d))
- **docx:** Refuse typing at the ends of a content-locked block ([3c49b40](https://github.com/ChristopherVR/ooxml/commit/3c49b40f7fcfae5a70015c01364aec2d60e3ec38))
- **geometry:** Draw the wedge callout presets as ECMA-376 defines them ([2a4f2b0](https://github.com/ChristopherVR/ooxml/commit/2a4f2b0b100040a5ab1352f6d36f6a5e64d82a58))
- **math:** Draw the minus sign and space only a binary one ([1620cb0](https://github.com/ChristopherVR/ooxml/commit/1620cb00486131c99437dc13259b5d82c5d7d63f))
- **geometry:** Keep the oval callout on its ellipse ([a98b010](https://github.com/ChristopherVR/ooxml/commit/a98b010ba274ae857d2bd3bc618761a1dce4bb39))
- **teams:** Keep message times strictly increasing per client ([972da6f](https://github.com/ChristopherVR/ooxml/commit/972da6f4bd662b72fd4eb8dacfa0cdf258694d92))
- **site:** Keep teams consent and formula rules in the new assistant ([7e43888](https://github.com/ChristopherVR/ooxml/commit/7e43888029e8ff1108a0e7b8fb9e7450d47b24bb))

### Refactor

- **pptx:** Drop the unused jszip import from the chart parser ([66f67dc](https://github.com/ChristopherVR/ooxml/commit/66f67dcb464a61015ea3b56fcc6578a29e7016a5))
- **chart:** Move waterfall steps and range into the chart area ([b6fa22a](https://github.com/ChristopherVR/ooxml/commit/b6fa22a987d40b63f3369868083435eba35afa83))
- **chart:** Move number formats, palettes, axis labels, auto title to core ([26bde03](https://github.com/ChristopherVR/ooxml/commit/26bde036faed3663e9a037b8e755e5c517202263))
- **chart:** Move bubble, stock, treemap and sunburst geometry to core ([a465fa6](https://github.com/ChristopherVR/ooxml/commit/a465fa63dfd1f15c67a22a97fa5ed9c97d2965e8))
- **pptx:** Read chart trendlines and error bars from core model ([fa53d97](https://github.com/ChristopherVR/ooxml/commit/fa53d97b442df87b60ac1be85f5e4480bb8fec7a))

### Testing

- **xlsx:** Guard the spill test cell lookups for the linter ([e1810b5](https://github.com/ChristopherVR/ooxml/commit/e1810b5d89d1153edd9da5991adcc929e3546e03))
- **xlsx:** Refresh authored PivotTables in Excel during acceptance ([9901e5b](https://github.com/ChristopherVR/ooxml/commit/9901e5bec604779ba1f5d83aa0e6ecda4adccb0f))
- **teams:** Check every sync room carries the signed session ([e6c386c](https://github.com/ChristopherVR/ooxml/commit/e6c386ca75852e30cba6c7657abbe725fe8b0edb))

### Chores

- **ui:** Drop an unused chart import and clarify the note lookup ([76ed6f1](https://github.com/ChristopherVR/ooxml/commit/76ed6f1bbd263129a0d3950119a8f062ba020d9f))

## [1.13.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.13.0) - 2026-10-15

### Features

- **teams:** Pin and unpin channel messages with a pinned bar ([00a4c60](https://github.com/ChristopherVR/ooxml/commit/00a4c603e4b95be5242ec30257a64e4355bc76c4))
- **teams:** Add do not disturb, status messages and idle away ([e0adea1](https://github.com/ChristopherVR/ooxml/commit/e0adea1e772fd89eaec14c87b4874823d62b0912))
- **xlsx:** Add bessel, odd-period, pivot and dynamic functions ([19cb829](https://github.com/ChristopherVR/ooxml/commit/19cb8290a24cf1283180dc125ce07d827876d5e3))
- **ui:** Lay out the shared options dialog like powerpoint's ([a999f52](https://github.com/ChristopherVR/ooxml/commit/a999f526af7aac2c4758e421e939ed0738b694f0))
- **ui:** Show excel's options in the shared options dialog ([c5db5ed](https://github.com/ChristopherVR/ooxml/commit/c5db5ed7cf41314bd2e2e5dfef5c9a8d12e1c2fe))

### Bug Fixes

- **ui:** Translate the ribbon scroll chevron label ([b894992](https://github.com/ChristopherVR/ooxml/commit/b894992014f0aa21f6922f8a55da6db2b12b122e))

## [1.12.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.12.2) - 2026-10-14

### Bug Fixes

- **ui:** Translate swatch, style and scheme names in pptx menus ([#83](https://github.com/ChristopherVR/ooxml/issues/83)) ([bd2908c](https://github.com/ChristopherVR/ooxml/commit/bd2908c03147486699378b99eccbba6cbd5b8efd))
- **ui:** Draw the table background and lines of Themed Styles ([#76](https://github.com/ChristopherVR/ooxml/issues/76)) ([8ea2382](https://github.com/ChristopherVR/ooxml/commit/8ea2382e8654a12e5f8f48d1658c8e872b7c9cc0))

### Documentation

- Fix dead docx-viewer and code of conduct links ([fc53a08](https://github.com/ChristopherVR/ooxml/commit/fc53a08ed6e7a0165eba76564026b8ebad5ea0f8))

## [1.12.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.12.1) - 2026-10-14

### Performance

- **docx:** Cache review data during caret updates ([e03f8c8](https://github.com/ChristopherVR/ooxml/commit/e03f8c82fdb7418e84bc0c041139e969c1633e17))

## [1.12.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.12.0) - 2026-10-14

### Features

- **docx:** Honor measured ordinary-cell paragraph keeping ([d4eba24](https://github.com/ChristopherVR/ooxml/commit/d4eba244e7a90bac3fb8350e39b34dc716a96499))
- **visio:** Count bounded non-fit print tiles ([df10a0a](https://github.com/ChristopherVR/ooxml/commit/df10a0a989a43c3478c9bb4844b358241874d708))
- **ui:** Align word excel and visio toolbar surfaces ([9decf56](https://github.com/ChristopherVR/ooxml/commit/9decf563e3c58be4fd7c767a4a459fad3a45029f))

### Bug Fixes

- **docx:** Refuse merged nested fragment profiles ([46d8b5a](https://github.com/ChristopherVR/ooxml/commit/46d8b5a922097a0a84a48b48d40acd50dc31bc83))
- **docx:** Retain zero row heights for bounded admission ([3ff4272](https://github.com/ChristopherVR/ooxml/commit/3ff42724e5577091af186b5fd7e1f783b1747b03))
- **visio:** Refuse unmeasured print fit-target metadata ([b59d87b](https://github.com/ChristopherVR/ooxml/commit/b59d87b87a40b1713ad4e754c11d28758da69d53))
- **ui:** Restore native touch dragging across office editors ([6dc56a1](https://github.com/ChristopherVR/ooxml/commit/6dc56a1ba920b230d2a2d23572378305e74fe48c))

### Performance

- **docx:** Reduce typing and print layout work ([5b06c49](https://github.com/ChristopherVR/ooxml/commit/5b06c49ece0ecb00bd24c40db9df6ac7777e3525))

### Testing

- **docx:** Measure border and margin startup fit boundaries ([5c87bd9](https://github.com/ChristopherVR/ooxml/commit/5c87bd99873000445bd3e5625af256f9d663e5ed))
- **visio:** Preserve native print-default precedence ([b9e027a](https://github.com/ChristopherVR/ooxml/commit/b9e027af9e4838f2fba8629edfef3c771442d48c))

## [1.11.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.11.0) - 2026-10-13

### Features

- **visio:** Edit existing local text fields atomically ([cd62ff3](https://github.com/ChristopherVR/ooxml/commit/cd62ff30574677c72820e0dcb9609afefedb5acf))
- **xlsx:** Author bounded pivot report filters ([14b2cfe](https://github.com/ChristopherVR/ooxml/commit/14b2cfe6c985c07ab9d53e67107f58fdf55b29ed))
- **xlsx:** Configure cached report filter selections ([993a5a5](https://github.com/ChristopherVR/ooxml/commit/993a5a5922656ccf90f8eece099a804200a7bc43))
- **xlsx:** Support nested pivot report filters ([550c3a4](https://github.com/ChristopherVR/ooxml/commit/550c3a4a948dcf8623e49c4c7c96c9a4b2e9ebed))

### Bug Fixes

- **pptx:** Honor bounded native visibility timing ([a9a70ee](https://github.com/ChristopherVR/ooxml/commit/a9a70ee8c87abfb32ab3d6999c633073bf5e3ddd))
- **visio:** Restrict existing field choices to local sources ([4b718d5](https://github.com/ChristopherVR/ooxml/commit/4b718d5d178485377044635f9f110d583aabdfe0))
- **docx:** Repeat outer borders on nested table fragments ([1cc842c](https://github.com/ChristopherVR/ooxml/commit/1cc842c520d1efedb702e2d3dc5298f246367904))
- **pptx:** Separate inherited paragraph spacing from direct values ([0c23aac](https://github.com/ChristopherVR/ooxml/commit/0c23aac22831cff4dd4a7dcce79d48ffdaab27b4))
- **docx:** Admit native three-row nested table fragments ([bb5145d](https://github.com/ChristopherVR/ooxml/commit/bb5145d755b34de1ede8c3b7d86f6ae50d583968))
- **docx:** Reflow automatic nested rows by structure ([c0c52ad](https://github.com/ChristopherVR/ooxml/commit/c0c52adbb287cfca5db8eb0a330a954d87e1c3ed))
- **xlsx:** Guard every repeated report column extension ([7a0f768](https://github.com/ChristopherVR/ooxml/commit/7a0f768be908da73f9211632532f26a1f80bb8f8))
- **ui:** Describe measured pivot report filter limits ([e3adfc8](https://github.com/ChristopherVR/ooxml/commit/e3adfc800a994de92d717bcc8cfa3d5aa62be0d4))
- **pptx:** Apply explicit line spacing over inherited paragraphs ([aad8c3d](https://github.com/ChristopherVR/ooxml/commit/aad8c3deba0eb8d4034b208e732136a5fc24cc92))

### Refactor

- **pptx:** Move native effect timing decisions into core ([61e030a](https://github.com/ChristopherVR/ooxml/commit/61e030a0ae01032d36da6211577bf340187052c8))
- **pptx:** Share font and whole-text edits in core ([94be138](https://github.com/ChristopherVR/ooxml/commit/94be138efc09b3eb881436eb07be54af0460fb45))
- **pptx:** Move paragraph and run style scope into core ([3ba3c32](https://github.com/ChristopherVR/ooxml/commit/3ba3c321af80efed2da4c6d58a377070c59b7c07))
- **pptx:** Move paragraph geometry cascade into core ([2dfc204](https://github.com/ChristopherVR/ooxml/commit/2dfc204f3ae6a84fed476073b3aefae261ab8d6b))

### Dependencies

- **deps:** Update production dependencies ([bbd7b53](https://github.com/ChristopherVR/ooxml/commit/bbd7b530c7a15ae48c0c0a2ff4ea7dcdfe2aa9d3))

## [1.10.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.10.0) - 2026-10-13

### Features

- **xlsx:** Author tabular pivots with multiple measures ([177a04d](https://github.com/ChristopherVR/ooxml/commit/177a04dce47fc620c5d54c32fbdc91a5188e73cc))
- **ui:** Configure multiple pivot value fields ([eede3cb](https://github.com/ChristopherVR/ooxml/commit/eede3cb1b058435cac114e716db834ffe5757685))
- **xlsx:** Support multiple measures in compact pivot layouts ([a58f7c1](https://github.com/ChristopherVR/ooxml/commit/a58f7c106f90f844e2ebe5497e9ece34ae956612))
- **ui:** Enable multi-measure pivot report layouts ([bf559ec](https://github.com/ChristopherVR/ooxml/commit/bf559ec8eb746c7b3a1d1869d04e0a4a51467c94))
- **xlsx:** Support tabular pivot values on rows ([7234296](https://github.com/ChristopherVR/ooxml/commit/72342960fccb8037decc6a411af5b70171a5a821))
- **ui:** Expose pivot values axis settings ([9a6164f](https://github.com/ChristopherVR/ooxml/commit/9a6164f7e0fa3ed23b337a601eba0524c4921323))
- **visio:** Match native general numeric field pictures ([ebf9f7a](https://github.com/ChristopherVR/ooxml/commit/ebf9f7aed842a097fe8d0d622b49846c5729dd48))
- **xlsx:** Support compact and outline pivot row values ([f9e5e70](https://github.com/ChristopherVR/ooxml/commit/f9e5e70bcc839049540f04a35dc5f6b7d60d0f0f))
- **ui:** Enable compact and outline pivot row values ([87535e3](https://github.com/ChristopherVR/ooxml/commit/87535e36cbf34ff04e0811d2db8ab70f15268580))
- **visio:** Format native fractional text field pictures ([ae6b922](https://github.com/ChristopherVR/ooxml/commit/ae6b9228565b246178a71b2ebd016e3736364db6))
- **pptx:** Support native whole-shape fixed opacity ([7b0372e](https://github.com/ChristopherVR/ooxml/commit/7b0372ed918da963bf662a302334e68ccae972e2))
- **visio:** Render explicit iso date field pictures ([ea3b0fa](https://github.com/ChristopherVR/ooxml/commit/ea3b0faf51ab3f37956cfbae7e5649555ee22ce4))
- **xlsx:** Support nested tabular values on rows ([a00cff7](https://github.com/ChristopherVR/ooxml/commit/a00cff761b229bbd4570eefa34dad291290dec10))
- **xlsx:** Author nested values rows in tabular reports ([fc94245](https://github.com/ChristopherVR/ooxml/commit/fc94245e3f5f2caf6529a2daabf4f0e25229c47e))
- **visio:** Offer iso dates in the field dialog ([53ed043](https://github.com/ChristopherVR/ooxml/commit/53ed043f3b2a46e7736f7389eabcbb5f409cd859))
- **xlsx:** Support nested compact and outline row values ([9a56995](https://github.com/ChristopherVR/ooxml/commit/9a56995bcbaa095984639b33034aa62fdb05ffad))
- **xlsx:** Configure nested row values in all report layouts ([f85cf7d](https://github.com/ChristopherVR/ooxml/commit/f85cf7d8c2580d5b4bc1b0fb8a0451efe577dc9a))
- **xlsx:** Place row values between two pivot fields ([a8ad158](https://github.com/ChristopherVR/ooxml/commit/a8ad158fd445a23eeef478e8b2f1d93c0e4e3ac7))
- **xlsx:** Configure between-field values placement ([07357bd](https://github.com/ChristopherVR/ooxml/commit/07357bde39824b7f433be2d6b7e7529b1ed041b0))
- **xlsx:** Place pivot values before two row fields ([3e77360](https://github.com/ChristopherVR/ooxml/commit/3e77360b8c63d7f82c874e5ce013bee93b55bd48))
- **xlsx:** Configure values before row fields ([4f33bd0](https://github.com/ChristopherVR/ooxml/commit/4f33bd0d2e27dd113ef21aeead69b8940a70b8bb))

### Bug Fixes

- **visio:** Match native decimal rounding and scientific display ([0c0f505](https://github.com/ChristopherVR/ooxml/commit/0c0f5054904f40bf76342767a5c708989323c152))
- **docx:** Paint nested table frames in print layout ([0415699](https://github.com/ChristopherVR/ooxml/commit/04156990666eb0a9f7b449c78266000578ea3797))
- **docx:** Align nested print text with cell frames ([66a9030](https://github.com/ChristopherVR/ooxml/commit/66a9030394475733f58f06541206274e27652c15))
- **xlsx:** Normalize Values headers during pivot conversion ([231be1f](https://github.com/ChristopherVR/ooxml/commit/231be1f2c3eb74521ede5e6f818753c15f3451e5))
- **xlsx:** Check pivot budgets before allocating headers ([ccfc390](https://github.com/ChristopherVR/ooxml/commit/ccfc39065ae692a1fe6bfef975545f50e4edf4f8))
- **docx:** Reserve measured nested table border insets ([d59a6ea](https://github.com/ChristopherVR/ooxml/commit/d59a6ea69bad31c395f58caed8160ebac082209f))
- **docx:** Measure parallel table continuation bottoms ([854baa9](https://github.com/ChristopherVR/ooxml/commit/854baa92664b927016425e8b1edaf7148d7d387b))
- **core:** Keep custom-geometry command order through undo and copies ([3ea38e4](https://github.com/ChristopherVR/ooxml/commit/3ea38e4e1bda2a59e7677b3c1783d53a89d1ab75))
- **docx:** Align asymmetric nested borders and paint layers ([7540cb1](https://github.com/ChristopherVR/ooxml/commit/7540cb1e6224298dec801134537512383aee007b))
- **pptx:** Bound legacy opacity image hint admission ([a37cf62](https://github.com/ChristopherVR/ooxml/commit/a37cf62761132b312825d3fe9d317a28317c7e85))
- **pptx:** Refuse unresolved opacity parent timing ([8d71f85](https://github.com/ChristopherVR/ooxml/commit/8d71f852f4f0f4305cf18dbafcb151f6c7fac280))
- **visio:** Render inherited native iso date fields ([2aaa106](https://github.com/ChristopherVR/ooxml/commit/2aaa1064bb88df629aa37299fbd0271c1857d040))
- **visio:** Preserve unmeasured short date locale caches ([0a8e568](https://github.com/ChristopherVR/ooxml/commit/0a8e568bc9d5c19054db6dfc95e856d029f5996f))
- **docx:** Reserve ordinary table border occupancy ([b7c904e](https://github.com/ChristopherVR/ooxml/commit/b7c904edc474c2704075a97ad029890c42b46d7c))
- **docx:** Continue nested footnote cells independently ([9317984](https://github.com/ChristopherVR/ooxml/commit/931798444bea75834fff1b6c3d04ddc1b0489529))
- **visio:** Write native date field and format caches ([81e07a7](https://github.com/ChristopherVR/ooxml/commit/81e07a7b82fbc10e2de04c21947efddaa919d345))
- **docx:** Gate note cuts on authoritative cell stories ([e52268d](https://github.com/ChristopherVR/ooxml/commit/e52268d975814e50130b6d30b29539f201193c97))
- **docx:** Bound independent note cuts to one reference ([752edb8](https://github.com/ChristopherVR/ooxml/commit/752edb8f988980e0edfa0a52591c74a42d65c4e3))
- **docx:** Release consumed nested footnote constraints ([00753c3](https://github.com/ChristopherVR/ooxml/commit/00753c3bb0bc287a1e47072a7c42965822ac8533))
- **pptx:** Preserve native after-state and honor explicit clears ([aed6a6d](https://github.com/ChristopherVR/ooxml/commit/aed6a6d03410a3487589edfa6f57bde243885a45))
- **docx:** Finish nested note tails in balanced columns ([c2cfc99](https://github.com/ChristopherVR/ooxml/commit/c2cfc99df37ae572d650bea9b50d64546199a967))
- **pptx:** Number slide-number fields from the deck's firstSlideNum ([ec9b18d](https://github.com/ChristopherVR/ooxml/commit/ec9b18da409c971d08491b612db228fa32608c8b))
- **pptx:** Keep a:br line breaks soft through an inline text edit ([02e40ae](https://github.com/ChristopherVR/ooxml/commit/02e40aed4dc8388676d5b852d73e5065d1b3846f))
- **pptx:** Drop cleared cell fills and borders when the style changes ([99dbeba](https://github.com/ChristopherVR/ooxml/commit/99dbebabe7fd885998c52360610781e10ad3bca4))
- **visio:** Refresh measured inherited system date fields ([9eaf6b8](https://github.com/ChristopherVR/ooxml/commit/9eaf6b86c522596ac1ce4ba9112cfa797de5b0a7))
- **visio:** Refresh validated inherited english long dates ([f80dd00](https://github.com/ChristopherVR/ooxml/commit/f80dd00cdf77ecc13bf6b56d63acd44552b31a91))

### Testing

- **visio:** Cover native fraction approximation boundaries ([b92ced3](https://github.com/ChristopherVR/ooxml/commit/b92ced30232d2c916394abfab15bd718c59c8e97))
- **docx:** Record native long interacting note continuations ([327efc9](https://github.com/ChristopherVR/ooxml/commit/327efc9c91f5bd2491a0ee71ed63dd6eb55be92a))

### Styling

- **visio:** Format native boundary regression ([8fd285d](https://github.com/ChristopherVR/ooxml/commit/8fd285d5a90976f7ef02e654b3f056b79d2120f5))

## [1.9.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.9.0) - 2026-10-13

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

- **xlsx:** Write native-compatible typed pivot caches ([a4ace48](https://github.com/ChristopherVR/ooxml/commit/a4ace48e4fb207971c60f570698d20c6f72bdc6f))
- **visio:** Match native field casing and curved walking glue ([7238edb](https://github.com/ChristopherVR/ooxml/commit/7238edb1e571320ae2e50f40946ce963a89a24a3))
- **docx:** Paginate footnotes referenced inside body tables ([e137736](https://github.com/ChristopherVR/ooxml/commit/e137736a5c11cc791f4a02da87e3f517205de1e4))
- **docx:** Retain cell padding when splitting table footnotes ([0e36444](https://github.com/ChristopherVR/ooxml/commit/0e36444a381f527d8cd03729e7bedd079ef3d29e))
- **pptx:** Preserve chart workbooks across repeated saves ([b684d66](https://github.com/ChristopherVR/ooxml/commit/b684d66f7da744b8d322178e9db0b5df05fcf0fa))
- **visio:** Match native corner curves and text placement ([0b50379](https://github.com/ChristopherVR/ooxml/commit/0b50379d172a1c8e50b8bd849c03d0713f31a6e5))
- **xlsx:** Preserve conflicting pivot layout definitions ([333c74d](https://github.com/ChristopherVR/ooxml/commit/333c74d403380c141becf9ccb3bc573d8c35e5b5))
- **docx:** Paginate nested table notes and guard source structure ([f67b8c3](https://github.com/ChristopherVR/ooxml/commit/f67b8c3491464230a3f8fd9dddf7654615d37af5))
- **visio:** Check supplied obstacles along corner curves ([59b3fa5](https://github.com/ChristopherVR/ooxml/commit/59b3fa52b867a7adff98bacd0ab08edefab10a53))
- **pptx:** Typecheck clipboard fixture without build output ([ee3769f](https://github.com/ChristopherVR/ooxml/commit/ee3769ffce40280cde48317020fb62b858e4a6ed))

### Testing

- **docx:** Remove unused round-trip test import ([9c913bd](https://github.com/ChristopherVR/ooxml/commit/9c913bdce9a207f063a48414658ab6844083c98d))

## [1.8.4](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.8.4) - 2026-10-12

### Bug Fixes

- **xlsx:** Rebind self references when copying sheets ([#55](https://github.com/ChristopherVR/ooxml/issues/55)) ([896fb0f](https://github.com/ChristopherVR/ooxml/commit/896fb0f3bcb5a3c6ac367464a8264e63256c040c))

## [1.8.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.8.3) - 2026-10-12

### Bug Fixes

- **xlsx:** Keep hidden records in place when sorting ([#53](https://github.com/ChristopherVR/ooxml/issues/53)) ([d073bec](https://github.com/ChristopherVR/ooxml/commit/d073bec4b4049f523893090fe069d94f28cc1014))

### Testing

- **pptx:** Isolate the raw-key inspector test and widen a perf margin ([e66aaec](https://github.com/ChristopherVR/ooxml/commit/e66aaec24adc0b68472eae1fc42aa4598e09c352))

## [1.8.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.8.2) - 2026-10-11

### Testing

- **ui:** Keep each DOM environment and test state in its own worker ([6d52bec](https://github.com/ChristopherVR/ooxml/commit/6d52bec41343129aa621e11e173fe3c706e82b6e))

### Styling

- Fix the oxfmt and oxlint warnings outside the viewers ([cd66f32](https://github.com/ChristopherVR/ooxml/commit/cd66f324a45aad1d9bd680d1518f0dbe8baf8fc3))

## [1.8.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.8.1) - 2026-10-10

### Bug Fixes

- **xlsx:** Reject insertions that discard boundary data ([#49](https://github.com/ChristopherVR/ooxml/issues/49)) ([ee4414d](https://github.com/ChristopherVR/ooxml/commit/ee4414d90569b64f940428108297a5096a6f355f))

### Testing

- **ci:** Share workers across core test files ([a3087e5](https://github.com/ChristopherVR/ooxml/commit/a3087e55b8ac838a6d379cdcd3d05d77870e7c7b))

## [1.8.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.8.0) - 2026-10-10

### Features

- **color:** Add the Office theme colour palette to the color area ([646b6b5](https://github.com/ChristopherVR/ooxml/commit/646b6b532a304013f0c8783801ab3401de8a1950))
- **visio:** Add AutoConnect arrows with a Quick Shapes mini toolbar ([55c5f4b](https://github.com/ChristopherVR/ooxml/commit/55c5f4bb7652d7f3ea70515357ab4815b8461b72))
- **visio:** Read the document stencil and drop masters as instances ([d42becb](https://github.com/ChristopherVR/ooxml/commit/d42becbf8a535bc31032f0b46302cbd5f0c75a50))
- **visio:** Route right-angle connectors around placeable shapes ([4acecba](https://github.com/ChristopherVR/ooxml/commit/4acecba8ed23fd56f769aa9398c9ce1997f0eade))
- **visio:** Resize, rotate and format stencil instances like Visio ([9065221](https://github.com/ChristopherVR/ooxml/commit/90652214da054224a7437ddcd05aa504577b9c24))
- **visio:** Offer resize, move, rotate and formatting on stencil shapes ([6795c0e](https://github.com/ChristopherVR/ooxml/commit/6795c0ed6106f781178a49aecefd687993998479))
- **ribbon:** Add a core area for the ribbon add-in tab descriptor ([7b2fa48](https://github.com/ChristopherVR/ooxml/commit/7b2fa48d2562d7483e03acf573e74404342b8f2a))
- **visio:** Make ribbon add-in tabs a prop and event of every binding ([6e6a157](https://github.com/ChristopherVR/ooxml/commit/6e6a157bbced977145d428e2ca01168e0a413deb))
- **visio:** Draw line jumps and toggle them from Design > Connectors ([842ddbb](https://github.com/ChristopherVR/ooxml/commit/842ddbb2f64602a3c201678d73bb01aa45cfd9d7))
- **visio:** Add the Ruler & Grid dialog, saved to the page ([f6646a8](https://github.com/ChristopherVR/ooxml/commit/f6646a84815af9148b7e18351eb4b2488b47d6d4))
- **visio:** Add the Snap & Glue dialog, saved to the drawing ([f96c927](https://github.com/ChristopherVR/ooxml/commit/f96c927906a6e4d6872b4667439726460d91dd61))
- **visio:** Insert a horizontal legend from Data > Insert Legend ([92ac505](https://github.com/ChristopherVR/ooxml/commit/92ac5055f1325d1bff1472d822d31d027d80455c))
- **visio:** Edit instances of group masters as Visio saves them ([9c75d99](https://github.com/ChristopherVR/ooxml/commit/9c75d99290521b8dfb7af625ff340612e25ee3e0))
- **visio:** Edit stencil groups and their parts in the viewer ([977bc39](https://github.com/ChristopherVR/ooxml/commit/977bc397b2255bfc8e10462dadfec63544a3816c))
- **text:** Add the fonts' design advance widths, regular and bold ([15f7eeb](https://github.com/ChristopherVR/ooxml/commit/15f7eebb692c4cf43d8febdd49070e583daef22a))
- **visio:** Size shapes from their text as Visio does ([e709f5f](https://github.com/ChristopherVR/ooxml/commit/e709f5fd055d76b50153f2271f5544b29ac05d02))
- **visio:** Lay Visio's Dynamic connector out again, as Visio saves it ([24cbff5](https://github.com/ChristopherVR/ooxml/commit/24cbff57daa30ef5b2d445ec24c36c6c5b36a681))
- **visio:** Offer glue and end handles on stencil shapes and connectors ([7d815b9](https://github.com/ChristopherVR/ooxml/commit/7d815b9f29508f5427a70abcab7728d9a57a785d))
- **visio:** Drop the Dynamic connector from the Document Stencil ([93eb00a](https://github.com/ChristopherVR/ooxml/commit/93eb00ae70b5ffd830c7644a1910fac39faf7b68))
- **visio:** Delete, copy, order, group and re-master stencil shapes ([d46f268](https://github.com/ChristopherVR/ooxml/commit/d46f2688d1498ba11ed6060a1c340e4843e3fa2a))
- **visio:** Save picked theme colours as Visio's theme formulas ([90995ba](https://github.com/ChristopherVR/ooxml/commit/90995ba8ad94e1b82f277a4c44405a3b77236a5f))
- **visio:** Pick theme colours that follow the page theme ([d29c9c7](https://github.com/ChristopherVR/ooxml/commit/d29c9c7b8ae343821db3461ca55af2283d6a19ef))
- **visio:** Edit line ends, line cap and corner rounding ([a856227](https://github.com/ChristopherVR/ooxml/commit/a85622756c98334e02653873769299a44a002f57))
- **visio:** Add stencil masters to a drawing and drop instances of them ([92474f8](https://github.com/ChristopherVR/ooxml/commit/92474f8007c1c542d524c0a4e952e8423052e4c7))
- **visio:** Build the sample drawing from stencil master instances ([1875739](https://github.com/ChristopherVR/ooxml/commit/1875739c832d19c78f687c2538db4a9d1b63aa5e))

### Bug Fixes

- **pptx:** Verify certificate chains with node:crypto, not node-forge ([51e6312](https://github.com/ChristopherVR/ooxml/commit/51e6312aa152a47ac11a6beced49b373adf0fb68))
- **visio:** Keep right-angle routes straight between close shapes ([9cdc70e](https://github.com/ChristopherVR/ooxml/commit/9cdc70e731bc50b580c90112eff53b41a4fd3c51))
- **visio:** Edit drawn shapes beside point glue and free connectors ([ba77be8](https://github.com/ChristopherVR/ooxml/commit/ba77be8bb626c0c5ae523a258f1966a1c87869c4))
- **visio:** Find master Scratch and Connections rows named from one ([e371a7c](https://github.com/ChristopherVR/ooxml/commit/e371a7cd5b16e8d827369e1c8ad65bcefc31b67b))
- **visio:** Keep a dropped shape's size cached beside its pin ([db60177](https://github.com/ChristopherVR/ooxml/commit/db601772e28ec9913cac6435139f3e469d7fcd34))
- **xlsx:** Rebase relative rules before deleting their anchors ([#47](https://github.com/ChristopherVR/ooxml/issues/47)) ([779768d](https://github.com/ChristopherVR/ooxml/commit/779768db3549cfd3c395aee779cabfa75e007feb))
- **xlsx:** Keep table filters on their source columns ([#51](https://github.com/ChristopherVR/ooxml/issues/51)) ([5368355](https://github.com/ChristopherVR/ooxml/commit/536835596eeb7ffee3908c8b75ea5f6db5191f42))
- **xlsx:** Decode windows-1252 CSV bytes without relying on the runtime ([4c371aa](https://github.com/ChristopherVR/ooxml/commit/4c371aad3f99eb98e9d478dec6e844dd1831e8df))

### Refactor

- **geometry:** Move the orthogonal connector router out of pptx ([7628894](https://github.com/ChristopherVR/ooxml/commit/7628894c847d098846b7250e8cf7cfee0b8e1ecd))
- **ui:** Move Excel and Word colour pickers onto the shared grid ([ab332b6](https://github.com/ChristopherVR/ooxml/commit/ab332b6ed7a4b80bc5016de05efb3fad694c85de))
- **visio:** Format date-picture tokens through a switch ([069774b](https://github.com/ChristopherVR/ooxml/commit/069774bdfdf476f52fef96032beb7995399e931b))

### Documentation

- **visio:** Format the Visio and ooxml-ui READMEs ([cd76ee2](https://github.com/ChristopherVR/ooxml/commit/cd76ee27e0961385949517e7e95a8405aecbcf35))

### Testing

- **visio:** Align stencil-instance tests with the merged behaviour ([1f184fd](https://github.com/ChristopherVR/ooxml/commit/1f184fde8172da3c5fee2d204b401a11e57633b7))
- **visio:** Drop an unused import from the master move tests ([15a50f9](https://github.com/ChristopherVR/ooxml/commit/15a50f997d72e3ae4990e368a071001ad0ae3e4b))
- **visio:** Prove the instance-based sample takes every shape command ([65e5427](https://github.com/ChristopherVR/ooxml/commit/65e5427b27c1e983807c2f060f571c3dc36be59a))

### Styling

- Apply oxfmt to files drifted from the formatter ([daeabf2](https://github.com/ChristopherVR/ooxml/commit/daeabf26125af866f90359ac8a5716234d5b8b36))

### Chores

- **ci:** Drop prettier leftovers and align oxfmt range ([6b85898](https://github.com/ChristopherVR/ooxml/commit/6b858982d1cbf2539cb523122d2ab048b6343ee8))

## [1.7.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.7.0) - 2026-10-10

### Features

- **visio:** Add createSampleVsdx, an editable sample drawing ([a8f4f7d](https://github.com/ChristopherVR/ooxml/commit/a8f4f7d1c378ce3a011b6638dedc8c8abb02acc1))
- **geometry:** Share arrow-key nudges and Shift angle snapping ([a839871](https://github.com/ChristopherVR/ooxml/commit/a8398713b97bcc9a193e6bd85aba51c55066eba8))
- **visio:** Nudge with arrow keys and snap moves to the grid ([efacb75](https://github.com/ChristopherVR/ooxml/commit/efacb75c896f64a7252b9200c5e0e0d4ba0f13e9))
- **visio:** Edit the text of stencil (master) instances ([672e65c](https://github.com/ChristopherVR/ooxml/commit/672e65cc1df75dc5cfbb8639938413c7d5e089c2))
- **visio:** Edit formatted, field and partial text of stencil shapes ([2779656](https://github.com/ChristopherVR/ooxml/commit/27796562ef728a6be884967e9ed84323c480c16b))
- **visio:** Add Layer Properties as Visio's dialog, saved to the drawing ([8091265](https://github.com/ChristopherVR/ooxml/commit/809126590cbae8e7e41c80e3261754fe2afbb0ee))

### Bug Fixes

- **visio:** Report refused edits in plain words ([41f6e2e](https://github.com/ChristopherVR/ooxml/commit/41f6e2e3cc3dbca4213f79530c802269e3da6603))

### Testing

- **core:** Fuzz the xml, formula, number format and hyperlink parsers ([c6911a0](https://github.com/ChristopherVR/ooxml/commit/c6911a00f79fb07eab294d885b760452b4bdaf72))

### Build & CI

- Scope write permissions to jobs and look up date tokens in a map ([c6021c4](https://github.com/ChristopherVR/ooxml/commit/c6021c46cb93d7531d7b45ec713c74fe81510104))

### Styling

- **visio:** Format the stencil-instance text test ([87608ae](https://github.com/ChristopherVR/ooxml/commit/87608aeb6c37839967137bbdc634fed643e7410b))

## [1.6.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.6.2) - 2026-10-09

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

### Chores

- **deps-dev:** Bump the development-minor-and-patch group across 1 directory with 6 updates ([#43](https://github.com/ChristopherVR/ooxml/issues/43)) ([e3a4861](https://github.com/ChristopherVR/ooxml/commit/e3a4861279b8d2bbf3b869617329eddb953b7838))

## [1.6.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.6.1) - 2026-10-09

### Bug Fixes

- **pptx:** Paint the outline gradient of chart line series ([#41](https://github.com/ChristopherVR/ooxml/issues/41)) ([158d44b](https://github.com/ChristopherVR/ooxml/commit/158d44bd79582591697e9d6ef8ccc5311dc6d1f4))

## [1.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.6.0) - 2026-10-09

### Features

- **visio:** Create freeform, pencil and arc paths ([eeb6956](https://github.com/ChristopherVR/ooxml/commit/eeb69560e18ce57e7e7dac4cfaf5ab8594525b55))
- **visio:** Change a local shape's outline to a basic shape ([8c808e9](https://github.com/ChristopherVR/ooxml/commit/8c808e962df0171d9d5057f18ec58af63b527ffa))
- **visio:** Add change case and format painter commands ([85440e1](https://github.com/ChristopherVR/ooxml/commit/85440e1b6a88608c297de91157648a057dda0102))
- **visio:** Add quick style and outer shadow format edits ([385f163](https://github.com/ChristopherVR/ooxml/commit/385f1635aa27b5f5fbfb591565b9ed599009efb2))
- **visio:** Insert pictures and edit shape links and screentips ([d17698c](https://github.com/ChristopherVR/ooxml/commit/d17698c6345dd5b873c0141e661c04ec07235cf9))
- **visio:** Add group-shapes and ungroup-shape edits ([1e905d7](https://github.com/ChristopherVR/ooxml/commit/1e905d7149631242ceb2909ca45b4701739d56a7))
- **visio:** Glue straight connectors to shapes with native dynamic glue ([e4970fa](https://github.com/ChristopherVR/ooxml/commit/e4970fa5cef4f8c6f3a131642e101a0141f4684d))
- **visio:** Add page setup, page properties and page decoration edits ([f473bd4](https://github.com/ChristopherVR/ooxml/commit/f473bd4804ade69b89e3b0aad9a5903c5735b680))
- **visio:** Keep auto size on standard sizes and match paper sizes ([5760f7b](https://github.com/ChristopherVR/ooxml/commit/5760f7ba03789026ccbb13dcd47e462c18053a55))
- **visio:** Add page themes and glow, soft edge and reflection effects ([709bd22](https://github.com/ChristopherVR/ooxml/commit/709bd2287e86a689eff8cb3b674f95f6da755d00))
- **visio:** Add comments, subprocess, diagram checks and shape reports ([6fc89d0](https://github.com/ChristopherVR/ooxml/commit/6fc89d019990be3baa2a71acac8fbe25a386f566))
- **visio:** Plan create new and create from selection subprocesses ([f52bbf1](https://github.com/ChristopherVR/ooxml/commit/f52bbf17e13df75a38c864d035d2d6ff594f92df))
- **visio:** Add stencil outlines, containers and callouts ([7a28511](https://github.com/ChristopherVR/ooxml/commit/7a28511f983a8c316968d8e01f5308894bc41118))
- **visio:** Edit shape data and save linked external data recordsets ([ca13e1d](https://github.com/ChristopherVR/ooxml/commit/ca13e1d184087b32cb03e8e5a727ec5fc0d362cf))
- **visio:** Add data import tables and data graphic edit planning ([350c37a](https://github.com/ChristopherVR/ooxml/commit/350c37a38e0a5809456aa02973d8ebfbb07c5030))
- **ui:** Enable the visio data tab with import, linking and data graphics ([fc88751](https://github.com/ChristopherVR/ooxml/commit/fc887514fc44310a895f4183fb903be415e951a6))
- **visio:** Add connection points, point glue and connector routes ([25c47f8](https://github.com/ChristopherVR/ooxml/commit/25c47f883bd89cdbe9f74799baea0ec0c9dee438))
- **visio:** Re-route glued connectors for move previews ([40e4dc5](https://github.com/ChristopherVR/ooxml/commit/40e4dc5d8276ae50531a872b86322bb0b9a47dc8))
- **visio:** Add re-layout, auto align, layer assignment and guide edits ([c4c9cb2](https://github.com/ChristopherVR/ooxml/commit/c4c9cb2dec0579c595fc0d6be0787fd1da4e8cd6))
- **visio:** Assign mixed layers in one transaction and paste text boxes ([3fedc6d](https://github.com/ChristopherVR/ooxml/commit/3fedc6dffe13ccaa1e6a100e33b90619b273ec23))
- **visio:** Add text dialog formatting, text block placement and text fields ([30f680b](https://github.com/ChristopherVR/ooxml/commit/30f680bacbb4783fdaddc62f25c6cfebc10b0f5d))
- **visio:** Add text dialog, text block and text draft helpers ([438a552](https://github.com/ChristopherVR/ooxml/commit/438a5529b051a601b7849b69fc3255e3d98ebd3d))

### Bug Fixes

- **visio:** Show one colour rule when every value is equal ([00bc35c](https://github.com/ChristopherVR/ooxml/commit/00bc35c94781e428bca99820aca216d7030cf73e))
- **pptx:** Save the size of empty table cell paragraphs ([5f73df1](https://github.com/ChristopherVR/ooxml/commit/5f73df19e7c77690d55cc1c33f1792d76d18851a))
- **visio:** Resolve code scanning alerts in connectors and links ([b370053](https://github.com/ChristopherVR/ooxml/commit/b370053ed6d6f96bb3ae9caeb84f3973a99ba9f8))

## [1.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.5.0) - 2026-10-09

### Features

- **visio:** Create every basic shapes outline ([10e2e82](https://github.com/ChristopherVR/ooxml/commit/10e2e825b041ff9dec6cf2555234f386c86ddfac))

### Bug Fixes

- **xlsx:** Clear stale filter row state ([#36](https://github.com/ChristopherVR/ooxml/issues/36)) ([158f6cc](https://github.com/ChristopherVR/ooxml/commit/158f6ccfb54b6965dedda325d2cc2df3adb187ef))

## [1.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.4.0) - 2026-10-08

### Features

- **visio:** Preserve rich replacements and set fixed page dimensions ([8882860](https://github.com/ChristopherVR/ooxml/commit/8882860df98992b16e6c0ac83666f7d69dab7bc8))

### Bug Fixes

- **ui:** Restore CI after the Visio find and replace changes ([f0ca01d](https://github.com/ChristopherVR/ooxml/commit/f0ca01d6618346cee74b723cfb9cb16f61eb1cb3))
- **xlsx:** Correct table, subtotal, save and sheet protection behavior ([#32](https://github.com/ChristopherVR/ooxml/issues/32)) ([af31ec6](https://github.com/ChristopherVR/ooxml/commit/af31ec670c229eabb8515e04bacf9f2c277ef053))
- **visio:** Resolve open code-scanning alerts ([a6bb188](https://github.com/ChristopherVR/ooxml/commit/a6bb188c4da9285a09589f6716b112aee82edb8b))

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [1.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.3.0) - 2026-10-08

### Features

- **visio:** Enable source-backed formatting and shape ordering ([c39f7be](https://github.com/ChristopherVR/ooxml/commit/c39f7bee397d27afa31155a96c01997122402773))
- **visio:** Enable multi-selection formatting and arrangement ([fb70098](https://github.com/ChristopherVR/ooxml/commit/fb70098a2b100d6ec37cd091acd618c91477cfb2))
- **visio:** Enable rich formatting and source duplication ([d50e388](https://github.com/ChristopherVR/ooxml/commit/d50e388f9190809d6a860894663cdbf774ca4dca))
- **visio:** Format proven inherited text cells ([cbb4b20](https://github.com/ChristopherVR/ooxml/commit/cbb4b20c66effb2f2f019a6f9c5ec7f83fa07608))
- **visio:** Enable source-backed shape clipboard ([2c13c1f](https://github.com/ChristopherVR/ooxml/commit/2c13c1f39998cf37e83d3949a6685f84f4b16e4d))
- **visio:** Delete reference-closed shape selections ([d02ed0e](https://github.com/ChristopherVR/ooxml/commit/d02ed0ebb1f97622436ef0944838eaefed9f7127))
- **visio:** Create editable blank drawing packages ([4bcb713](https://github.com/ChristopherVR/ooxml/commit/4bcb71385e260240bbd03e1349fc4ff020327bd7))
- **visio:** Enable new drawings and pointer selection gestures ([75ac409](https://github.com/ChristopherVR/ooxml/commit/75ac409e5e67c23203c765d3aa472b6817853090))
- **visio:** Add source-proven anchored shape resizing ([cff07b8](https://github.com/ChristopherVR/ooxml/commit/cff07b8d9d7521819f2f4c3c514d77ca77d1638c))
- **visio:** Enable resize handles and size and position pane ([7601406](https://github.com/ChristopherVR/ooxml/commit/76014064de44defbf80208cd0fab1f33ef17c00b))
- **visio:** Add text boxes and preserve logical text paragraphs ([cbc47ff](https://github.com/ChristopherVR/ooxml/commit/cbc47ffc77cd8145c5e6b633a3dd87a105ed4072))
- **visio:** Add native paint patterns and transparency edits ([3a9ba13](https://github.com/ChristopherVR/ooxml/commit/3a9ba136f30d71b030af7cb86eecceed1ce49e3d))
- **visio:** Add scoped replacement plans and SVG hairlines ([79fb8e8](https://github.com/ChristopherVR/ooxml/commit/79fb8e8f6b572a3b7d53d8ba660626f7d49a7e9d))

### Bug Fixes

- **pptx:** Write cell properties into an empty a:tcPr ([22e8a3b](https://github.com/ChristopherVR/ooxml/commit/22e8a3b9661b4dada517df5d24d47e8eb3fa94e9))
- **xlsx:** Read formula sheet prefixes without a backtracking regex ([4a19c87](https://github.com/ChristopherVR/ooxml/commit/4a19c87ce84643af7af28e22abb7917c7efc9771))

## [1.2.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.2.1) - 2026-10-08

### Bug Fixes

- **pptx:** Keep unedited empty table cells free of a:tcPr ([01d4f3d](https://github.com/ChristopherVR/ooxml/commit/01d4f3d36d5acd4fb795b4f33018efa0e3482b06))
- **pptx:** Default an absent c:gapWidth to 150 ([e16f565](https://github.com/ChristopherVR/ooxml/commit/e16f5651541cfbba9b7a1280e58463c041e507d9))
- **xlsx:** Format selection statistics with the cell number format ([649f556](https://github.com/ChristopherVR/ooxml/commit/649f5568b348083aa548b2e819da02435d68d64a))
- **xlsx:** Make the formula number lexeme linear ([9b32d49](https://github.com/ChristopherVR/ooxml/commit/9b32d493d49dde1ba9002da056c4a24a8885af11))

### Performance

- **xlsx:** Prepare the formula graph incrementally ([15dc223](https://github.com/ChristopherVR/ooxml/commit/15dc22334dc6a5561817b24399a4cb0675dae3b4))
- **xlsx:** Shift the formula graph on row and column edits ([e3ef7d8](https://github.com/ChristopherVR/ooxml/commit/e3ef7d8bd7cb29d0c8cdcc6bf8dee29c07cace38))
- **xlsx:** Rename sheet references without rebuilding the graph ([57094bf](https://github.com/ChristopherVR/ooxml/commit/57094bfadb0831e52d08553f7927176c66062a68))
- **xlsx:** Cache formula templates for structural rewrites ([6d477d0](https://github.com/ChristopherVR/ooxml/commit/6d477d0023ceede31a50326f6fd3761bb8d97284))

### Refactor

- **pptx:** Share the chart XML presence check for a:noFill ([ad7bae0](https://github.com/ChristopherVR/ooxml/commit/ad7bae0ea91758a2e7dcb581c2f472da394e5eb4))

### Testing

- **xlsx:** Structural edits while the formula graph is prepared ([67e0de5](https://github.com/ChristopherVR/ooxml/commit/67e0de5f621a1fdb3149a7a261f0d33b9e4073f4))

### Chores

- **xlsx:** Profile structural edits ([b3330b4](https://github.com/ChristopherVR/ooxml/commit/b3330b419664f5fa7494bc1deada2ba8e25cc232))

## [1.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.2.0) - 2026-10-08

### Features

- **chart:** Add the neutral chartSpace model ([2544c31](https://github.com/ChristopherVR/ooxml/commit/2544c31319b5cc723b26a2c37b519719693f62b8))
- **chart:** Parse chart parts from the shared DOM ([49c1b76](https://github.com/ChristopherVR/ooxml/commit/49c1b76e41b5fac972c52a5d8169e9ad14478cfd))
- **docx:** Load chart parts into the document model ([01ce179](https://github.com/ChristopherVR/ooxml/commit/01ce17985b5d5bd0399497af6fe370ae99413a53))
- **opc:** Patch property parts in place and model app statistics ([66c4cf1](https://github.com/ChristopherVR/ooxml/commit/66c4cf1043d73a6df05131d813cbae7f5af6e35e))
- **xlsx:** Model and read sparkline groups ([681f241](https://github.com/ChristopherVR/ooxml/commit/681f2418a0300ffb19287b57928544f4e90b0a95))
- **xlsx:** Lay out sparklines for painting ([d95575a](https://github.com/ChristopherVR/ooxml/commit/d95575ab9d67bc0c96c33181262b3de232214493))
- **xlsx:** Add the Yjs workbook adapter ([6289494](https://github.com/ChristopherVR/ooxml/commit/628949485ba4f14ffda1f131cf6c8b44fc5596bf))
- **xlsx:** Bind edit sessions to shared workbooks ([cd28ab3](https://github.com/ChristopherVR/ooxml/commit/cd28ab35755790f149632bd2dae0ad2a37d5935d))
- **chart:** Write chartSpace parts from the neutral model ([bf6920f](https://github.com/ChristopherVR/ooxml/commit/bf6920fb553ff32c32c3edf9bff652cad3bd71ac))
- **chart:** Accept text bodies given as XML and drop invalid characters ([816830f](https://github.com/ChristopherVR/ooxml/commit/816830fa29a42491d72b835a9b6eb7b5a617f234))
- **xlsx:** Lay out remote selections ([3770bb8](https://github.com/ChristopherVR/ooxml/commit/3770bb85d8d7106fcf67c822533c189e5087e1c1))
- **xlsx:** Let the editor context own the undo history ([46641d7](https://github.com/ChristopherVR/ooxml/commit/46641d76777a94c25897c391c8066d3b057f8212))
- **chart:** Render chartSpace parts to svg ([eb2c12f](https://github.com/ChristopherVR/ooxml/commit/eb2c12f877c0f4c7521a9d71e323120fb9084ca6))
- **docx:** Draw charts in the editor and print layout ([da2ed18](https://github.com/ChristopherVR/ooxml/commit/da2ed1850715cf85d01ecc4f3967aa48d9414348))
- **chart:** Model display units, data tables and walls ([b95531a](https://github.com/ChristopherVR/ooxml/commit/b95531a3ff55f09c83676291c4b4857d1943189c))
- **opc:** Support shared suite files and embedded editing ([41614b1](https://github.com/ChristopherVR/ooxml/commit/41614b10ef62522fe2dceae695b478b53cca293a))

### Bug Fixes

- **pptx:** Keep pptx/ui in the relaxed project the bundlers read ([fb2a0d3](https://github.com/ChristopherVR/ooxml/commit/fb2a0d3224cc8d524baae359de5098003332ef0e))
- **chart:** Paint coincident linear stops as the native one-step ramp ([eabbc5e](https://github.com/ChristopherVR/ooxml/commit/eabbc5e3dc968f0d40246293238ef88ae7b004f4))
- **pptx:** List line breaks in slide titles as spaces in app.xml ([45f496a](https://github.com/ChristopherVR/ooxml/commit/45f496ab747f90f7b8764a295c8d0df8414141c0))
- **pptx:** Export the neutral SmartArt type names from the public entry ([0f081d2](https://github.com/ChristopherVR/ooxml/commit/0f081d260fbaf034129114fcba6137c900dbf846))
- **xlsx:** Report sparklines as shown in the feature notes ([1db5daa](https://github.com/ChristopherVR/ooxml/commit/1db5daacd00458cdd74e6913ea5b6c627bbe6543))
- **chart:** Read the line chart group's own c:smooth ([a5cfd0d](https://github.com/ChristopherVR/ooxml/commit/a5cfd0d45af11e2b2214a9abdc86ce33ba5312be))
- **pptx:** Read chart style numbers from the mc:Fallback ([045faf7](https://github.com/ChristopherVR/ooxml/commit/045faf71cadf05f1c9bc4b3155003abe9520b764))
- **pptx:** Treat a bare data-label switch as on ([637f938](https://github.com/ChristopherVR/ooxml/commit/637f938814baced0d2f72aeca7142e2281db0ca3))
- **pptx:** Declare c15 and c16 prefixes in chart extension writers ([7b93a01](https://github.com/ChristopherVR/ooxml/commit/7b93a010a77c1ef8433f75fc8d2f3ad810659c75))
- **pptx:** Repair chart parts this library once saved with undeclared prefixes ([a5f77af](https://github.com/ChristopherVR/ooxml/commit/a5f77af3bf0a2661af35091da920ca707e4fcb96))
- **xlsx:** Adopt the shared workbook before seeding a room ([3b3aa8e](https://github.com/ChristopherVR/ooxml/commit/3b3aa8e3e4b5d54c062b2d955b0f1008ee9cebc0))
- **pptx:** Build chart style palettes over the deck theme ([7ab9527](https://github.com/ChristopherVR/ooxml/commit/7ab9527743dbb1dc1438d2d3e6e73553b13e54b7))
- **pptx:** Read each table cell paragraph's own layout ([bcc61f3](https://github.com/ChristopherVR/ooxml/commit/bcc61f3b3e98564e50d018a4a37576f3b237fb57))
- **pptx:** Save each table cell paragraph's own alignment ([0d43f35](https://github.com/ChristopherVR/ooxml/commit/0d43f3530d5cefd52b98193be98329b4ec66547b))
- **pptx:** Align every paragraph when a table cell is aligned ([1bb7e7d](https://github.com/ChristopherVR/ooxml/commit/1bb7e7de568a0d8a6808862fd6dc4367782e9686))
- **pptx:** Skip chart axis lines and gridlines set to no line ([c6a9fcb](https://github.com/ChristopherVR/ooxml/commit/c6a9fcb0fb07498acf6db64d9ee9417339a6f2ac))
- **pptx:** Apply a table cell run's character spacing ([6d1ee14](https://github.com/ChristopherVR/ooxml/commit/6d1ee148cbd0ea058ae032a7a3b8fe9cdace1ce1))
- **pptx:** Size an empty table cell from its end paragraph properties ([30da656](https://github.com/ChristopherVR/ooxml/commit/30da656aa60039fff3c912c20cd286dc49055adf))
- **teams:** Bound the content pattern ([78dde2d](https://github.com/ChristopherVR/ooxml/commit/78dde2dbd5f7161febe60f1dcd84586fe7668b4d))
- **xlsx:** Guard chart series edits against prototype keys ([c49e00f](https://github.com/ChristopherVR/ooxml/commit/c49e00f90bc48b2443841bdfa5d21965175496a8))

### Performance

- **xml:** Add a lightweight reader for large plain fragments ([5e0318f](https://github.com/ChristopherVR/ooxml/commit/5e0318facba5e0766dd6d1f26172eb94177fe471))
- **xlsx:** Read worksheet cell data without building a DOM ([d5addd2](https://github.com/ChristopherVR/ooxml/commit/d5addd2352df2b2aead2920136ae38775663bec2))
- **xlsx:** Trust stored values on the first recalculation ([a584a55](https://github.com/ChristopherVR/ooxml/commit/a584a550628a9725f2a637a671883e28d8ca1dc3))
- **xlsx:** Tokenize formulas without slicing the source ([a7ec241](https://github.com/ChristopherVR/ooxml/commit/a7ec241ea393436f4c27d4e7024eef1878460c82))

### Refactor

- **pptx:** Compile the ui helpers under the strict project ([edbff0a](https://github.com/ChristopherVR/ooxml/commit/edbff0a3fd1863efa394da6c9fe826514fda5f7c))
- **visio:** Read the theme through drawingml ([5281182](https://github.com/ChristopherVR/ooxml/commit/528118204a9405f3e8f6dd216fb9d7bdd0c62f93))
- **chart:** Use the Drawing type names ([76bf49c](https://github.com/ChristopherVR/ooxml/commit/76bf49c0a96519816c79375eed1aa9ead84f8b10))
- **xlsx:** Use the Drawing type names ([ab6fc07](https://github.com/ChristopherVR/ooxml/commit/ab6fc07bbf978d27b4f3335676331016a2bf736d))
- **xlsx:** Build chart objects from the shared chart model ([d0a660e](https://github.com/ChristopherVR/ooxml/commit/d0a660e6e54d20ed788594476360cbe8b3db4b9e))
- **diagram:** Add the neutral SmartArt model types ([b8f4e74](https://github.com/ChristopherVR/ooxml/commit/b8f4e7406362570e085c701dd1e342d534fcc048))
- **geometry:** Move the preset shape evaluator out of pptx ([c670910](https://github.com/ChristopherVR/ooxml/commit/c670910349b7f6d6d02c335a551518265efde6a9))
- **diagram:** Move the layout engine out of pptx ([91767c0](https://github.com/ChristopherVR/ooxml/commit/91767c03d4ec9fb2225c47eddd0adc850f187022))
- **pptx:** Update document properties through opc/properties ([a295f6a](https://github.com/ChristopherVR/ooxml/commit/a295f6a2182eb344e6c9e9a6fe70ba5366268b9c))
- **diagram:** Port the dgm:choose walkers to ordered XML ([2dfc037](https://github.com/ChristopherVR/ooxml/commit/2dfc0371bb0d4b72e5a3ef4b4cb71bd1a5cb7f81))
- **diagram:** Move the SmartArt constraint solver into diagram ([bd91ad8](https://github.com/ChristopherVR/ooxml/commit/bd91ad884cf00947957b9620309c885df21d6985))
- **diagram:** Move the SmartArt interpreter model into diagram ([b24d320](https://github.com/ChristopherVR/ooxml/commit/b24d3200af42eda7c1d7fcc7dcf592202e2fc0f8))
- **xlsx:** Write new charts through the shared writer ([2f898a7](https://github.com/ChristopherVR/ooxml/commit/2f898a7c560b3141df55b3544691fe5deefc118f))
- **chart:** Move the svg painter into the chart area ([0f9169d](https://github.com/ChristopherVR/ooxml/commit/0f9169d11c5bf92342bc291fdb3f8ef8d84927c5))
- **pptx:** Build chart models from the neutral parser ([1e38985](https://github.com/ChristopherVR/ooxml/commit/1e389856ae308843d0d86c83b41f6a62284ef1d2))
- **pptx:** Delete the legacy object-tree chart parser ([b4f0f4f](https://github.com/ChristopherVR/ooxml/commit/b4f0f4f8af4b2eefe49749f5b3c1b1ee7492dba9))
- **diagram:** Move the SmartArt hierarchy arranger into diagram ([989744c](https://github.com/ChristopherVR/ooxml/commit/989744c5a2ed121eda9b2b11a37815e295c14740))
- **diagram:** Move the SmartArt layout interpreters into diagram ([3725b73](https://github.com/ChristopherVR/ooxml/commit/3725b7359b5d845012fa7b84b3cddc421446cfbd))

### Documentation

- Describe the xlsx collaboration adapter ([d17cef3](https://github.com/ChristopherVR/ooxml/commit/d17cef34d8cd04f46419dcaf9989a99688b7fecf))

### Testing

- **pptx:** Compare chart parts with the neutral parser on every deck ([58b16be](https://github.com/ChristopherVR/ooxml/commit/58b16be72bdcf27e291ff5b9bebc139db9cd3607))

### Chores

- **xlsx:** Add a load profiler script ([60650a6](https://github.com/ChristopherVR/ooxml/commit/60650a6add08555a3c475b39139d88f36b773aa9))
- **scripts:** Expand every wildcard in package smoke subpaths ([f240a06](https://github.com/ChristopherVR/ooxml/commit/f240a06302ce04102f52a4df11dae529009dbc3b))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.1.0) - 2026-10-08

### Features

- **ui:** Enable proven visio group rotation interactions ([be9ca55](https://github.com/ChristopherVR/ooxml/commit/be9ca5586434432f1673d42f5a0a6560ace25d83))
- **docx:** Edit imported list instance starts ([fe839f1](https://github.com/ChristopherVR/ooxml/commit/fe839f108009479d89e105a74a25f4bd8bea6bff))
- **docx:** Track plain paragraph splits and joins ([0b11ba1](https://github.com/ChristopherVR/ooxml/commit/0b11ba1e3599cd082d68862010c4c2576a5b8410))
- **xlsx:** Render imported manual legend layouts ([5173750](https://github.com/ChristopherVR/ooxml/commit/5173750a97a560e50d408d3f2f39df6d162d78d5))

### Bug Fixes

- **docx:** Track complete simple field replacements ([e4a7ccd](https://github.com/ChristopherVR/ooxml/commit/e4a7ccdb6a0eea969ed9fc4b878d6114ed749425))
- **docx:** Preserve field lock and dirty state ([818d8cd](https://github.com/ChristopherVR/ooxml/commit/818d8cd2e0ff0f2ab14b7da49b211f296d43cec6))
- **docx:** Retain locked caches during field updates ([9cad7c9](https://github.com/ChristopherVR/ooxml/commit/9cad7c9b61fd76d900cef945f6766547e164e64e))
- **docx:** Repair complex field lock metadata after edits ([93b8b4e](https://github.com/ChristopherVR/ooxml/commit/93b8b4e9662c5aaf0bb7a94f599492d0441ffa93))
- **docx:** Exclude bookmarked tracked paragraph splits ([fcd6cb7](https://github.com/ChristopherVR/ooxml/commit/fcd6cb7caec0596511268f8f0799d518c9331cf1))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@1.0.0) - 2026-10-07

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
- **docx:** Resolve inline object formatting history ([9689395](https://github.com/ChristopherVR/ooxml/commit/9689395a265d78a0f472f4880b6191130d2a9b43))
- **docx:** Project prior inline object formatting ([73691b8](https://github.com/ChristopherVR/ooxml/commit/73691b88ba16cf53b10aaaf68696ecb1c95980e8))
- **visio:** Translate local straight lines through shared edits ([856a0e1](https://github.com/ChristopherVR/ooxml/commit/856a0e11344265832906831b5eaa236a5c4f1e80))
- **visio:** Delete local lines and honor endpoint locks ([92e9a2b](https://github.com/ChristopherVR/ooxml/commit/92e9a2b5be3890d97706e9660a1a867c4596cbbd))
- **docx:** Record inline object formatting in shared history ([5de353d](https://github.com/ChristopherVR/ooxml/commit/5de353df2cf0b61b921dd2fef0141485292db1d8))
- **xlsx:** Add native named gradient preset gallery ([47ec692](https://github.com/ChristopherVR/ooxml/commit/47ec69208803df71d035ba191653134ec475ff26))
- **docx:** Preserve hard-break formatting and shared history ([a4dfbc4](https://github.com/ChristopherVR/ooxml/commit/a4dfbc40088ad577195695092149bdd195d798f2))
- **visio:** Resize native local lines through the width cell ([cdb9b5f](https://github.com/ChristopherVR/ooxml/commit/cdb9b5f5e2f393de2e740489920e70bdc820f1c2))
- **docx:** Share font commands across text and inline objects ([469c0d4](https://github.com/ChristopherVR/ooxml/commit/469c0d4b8f9d3a43dbcb2ef907ce472c2ba8e5ad))
- **visio:** Edit native local line endpoints ([3822b31](https://github.com/ChristopherVR/ooxml/commit/3822b31f8a4a7b08945772200448ac2b296c1d02))
- **xlsx:** Add rectangular gradient geometry controls ([e10c21a](https://github.com/ChristopherVR/ooxml/commit/e10c21a08bf6fdc3238504d7434db3f0b2a8b68f))
- **visio:** Drag local line endpoints on the canvas ([f68d353](https://github.com/ChristopherVR/ooxml/commit/f68d35354731f13c283032d81d2d11aec9c4b783))
- **docx:** Merge independent inline properties in yjs ([988a721](https://github.com/ChristopherVR/ooxml/commit/988a721111607e6588b628de48745e5bfa1cef6b))
- **docx:** Preserve inline hyperlinks through shared commands ([d1e7bf0](https://github.com/ChristopherVR/ooxml/commit/d1e7bf05f53d7b84fa2b7da3aec4ea006f629b8f))
- **xlsx:** Author radial and path fills on chart series ([2d3652a](https://github.com/ChristopherVR/ooxml/commit/2d3652ad10b2422302dbe457bb6899f2e56cb7bb))
- **visio:** Create straight lines through shared drawing tools ([f67c10f](https://github.com/ChristopherVR/ooxml/commit/f67c10f3032f03e4c4b5b01f1caafe0395bb7af6))
- **docx:** Share independent inline comment anchors ([4397b4b](https://github.com/ChristopherVR/ooxml/commit/4397b4bf5629a7a9da21a34b286f01d2b0853136))
- **ui:** Refine shared files commands and sorting ([9372d61](https://github.com/ChristopherVR/ooxml/commit/9372d61fecf826888d184a48d3f3f1c063328174))
- **xlsx:** Format chart and plot background fills ([9237f7f](https://github.com/ChristopherVR/ooxml/commit/9237f7f2ec5851088a28d2af5ca82d02806f5438))
- **visio:** Create and resize native ellipses with shared logic ([6c4d008](https://github.com/ChristopherVR/ooxml/commit/6c4d00839469577dfff31a61e27150c7bd281c53))
- **xlsx:** Format and paint chart title and legend fills ([4255dcd](https://github.com/ChristopherVR/ooxml/commit/4255dcdee2fbf4c62ea9260ffcf3a9856ebb859b))
- **visio:** Rotate local shapes through shared editing ([c68e4f4](https://github.com/ChristopherVR/ooxml/commit/c68e4f471d9fdd2927bdb47a13c95b78c0bd706e))
- **visio:** Rotate shapes with shared pointer geometry ([407fa61](https://github.com/ChristopherVR/ooxml/commit/407fa614c7f919052d44d9e31f7abfbcfe6df50d))
- **visio:** Render live rotation previews through shared svg ([87c3957](https://github.com/ChristopherVR/ooxml/commit/87c3957b5967466db44cbc95b08f5cfff21d81a7))
- **visio:** Share native quarter-turn command preparation ([c342e79](https://github.com/ChristopherVR/ooxml/commit/c342e79b82930797a7bf3a16b89a3bbe86d0128e))
- **xlsx:** Render and preserve mixed chart title text ([4f3f87e](https://github.com/ChristopherVR/ooxml/commit/4f3f87ea6eff2802393f24448a1b48931282d012))
- **visio:** Flip local shapes through shared geometry edits ([cffd415](https://github.com/ChristopherVR/ooxml/commit/cffd415a9fd7a8866f3f07dae027850dccb17c82))
- **geometry:** Support proven visio flip formulas ([f4d411d](https://github.com/ChristopherVR/ooxml/commit/f4d411d1fc8b9c7fd6876f581e25c7aa979ed563))
- **xlsx:** Render and preserve chart title paragraph spacing ([4f0d619](https://github.com/ChristopherVR/ooxml/commit/4f0d6198d90ab3befbfa6b9806965f01492bce45))
- **geometry:** Reuse source proof for visio rotation ([5f90540](https://github.com/ChristopherVR/ooxml/commit/5f905400b9e364e7061c1d30d6b6a6318023df25))
- **xlsx:** Wrap chart titles with shared text flow ([14a33ed](https://github.com/ChristopherVR/ooxml/commit/14a33ed2ee58f8469e5d24a0a36923a3922f23c4))
- **docx:** Render native hexadecimal list labels ([a7cc8aa](https://github.com/ChristopherVR/ooxml/commit/a7cc8aac044b2ebc24294d1df6c3ae7d7893a406))
- **geometry:** Support proven local visio group rotation ([cd7e429](https://github.com/ChristopherVR/ooxml/commit/cd7e429a21f0e0d986ce2d6b6e6c0cced1afd1e8))

### Bug Fixes

- **visio:** Project oblique stroke gradients in physical bounds ([2a05f37](https://github.com/ChristopherVR/ooxml/commit/2a05f37fb9e7e2e7683d0566a2831cc0122360cf))
- **visio:** Project saved oblique fills through physical bounds ([938ceae](https://github.com/ChristopherVR/ooxml/commit/938ceae0aaa1a7db365f577f36550863757f3f8f))
- **visio:** Render gradients on height-zero line shapes ([2980419](https://github.com/ChristopherVR/ooxml/commit/2980419d1b67a0cf8fd00d4b1657fdb5a2aa89a4))
- **chart:** Match native scaled linear gradient paint ([8bd1171](https://github.com/ChristopherVR/ooxml/commit/8bd1171809210e886f9a459a5d634b2109ad06e9))
- **chart:** Paint native rectangular path gradients ([af81015](https://github.com/ChristopherVR/ooxml/commit/af81015bca6823dc505c26c006b2a99eb2e85c8f))
- **xlsx:** Paint native circular chart backgrounds ([e2bddbf](https://github.com/ChristopherVR/ooxml/commit/e2bddbfdafdfe31792ec9f9e6e5002350e5d2f4f))
- **docx:** Locate imported inline comment anchors ([7ecbe15](https://github.com/ChristopherVR/ooxml/commit/7ecbe15f2e9b25401316579ad35e855b5add6b90))
- **xlsx:** Preserve chart axis visibility on export ([da056bb](https://github.com/ChristopherVR/ooxml/commit/da056bb9ebbc70561d63b1dc9e1f558997c0f946))
- **xlsx:** Preserve direct chart fills on export ([c3b3491](https://github.com/ChristopherVR/ooxml/commit/c3b3491a109720a2e38599b3b5272e9d84964309))
- **visio:** Preserve drawn line endpoints beyond paper edges ([6cc3708](https://github.com/ChristopherVR/ooxml/commit/6cc37087579742326cacec4b8b670b869c1ad48f))
- **docx:** Anchor comments to complete complex fields ([065cb5c](https://github.com/ChristopherVR/ooxml/commit/065cb5c1c8554e9cdbe7b61a807d9efff1ffc2de))
- **visio:** Apply native drawing defaults to new rectangles ([13d6830](https://github.com/ChristopherVR/ooxml/commit/13d6830d08e9022e4d109777063b9de77b686a7b))
- **docx:** Undo comment records with their anchors ([c736fab](https://github.com/ChristopherVR/ooxml/commit/c736fabf6d76c7e5cf87f137745e7dc1ea21bdba))
- **docx:** Retain body edits during yjs settings updates ([62d3fbd](https://github.com/ChristopherVR/ooxml/commit/62d3fbd3a5b852cb78619f3006ec982e5bbc4230))
- **docx:** Retain complete simple field comment scopes ([30500ef](https://github.com/ChristopherVR/ooxml/commit/30500efc737d69b920530893bf4de24b629b1dd9))
- **visio:** Preserve drawn boxes beyond scaled paper edges ([0048416](https://github.com/ChristopherVR/ooxml/commit/0048416f0cca19f7d3888e4c71c1d1d09583792e))
- **docx:** Retain adjacent simple field boundaries ([9303f68](https://github.com/ChristopherVR/ooxml/commit/9303f68aab372012d833669aa4eb45ceffba7726))
- **docx:** Copy field results as formatted literal text ([797db1b](https://github.com/ChristopherVR/ooxml/commit/797db1ba8f8942b2b364833e6a065504486ceaac))
- **pptx:** Copy owned package parts when duplicating slides ([6940f93](https://github.com/ChristopherVR/ooxml/commit/6940f936cc1f804a498aea941480357ad4236edb))
- **pptx:** Generate complete native 3d chart axes ([6de35d9](https://github.com/ChristopherVR/ooxml/commit/6de35d96c68b4c7a6118b7bcb2a9ac27fa7f3ecc))
- **pptx:** Write text properties before paragraph content ([a5fe6e7](https://github.com/ChristopherVR/ooxml/commit/a5fe6e78fe4889804f4a553dcbf4ce4a1c6c92e4))
- **docx:** Retain simple fields during result replacements ([908c45e](https://github.com/ChristopherVR/ooxml/commit/908c45ea9da8bf6ca75b7a4faefab87db44711d4))
- **xlsx:** Preserve native built-in chart text defaults ([f3b9335](https://github.com/ChristopherVR/ooxml/commit/f3b93358bc0fd123c1eb83cfba9237e183ae31e6))
- **docx:** Preserve empty fields through deletion and review ([e3c06a7](https://github.com/ChristopherVR/ooxml/commit/e3c06a73632b21645ff6553ed85d5426bfe49e53))
- **docx:** Share field-aware deletion across clipboard paths ([408ecb8](https://github.com/ChristopherVR/ooxml/commit/408ecb8a12549b953777c3037074280110ddd017))
- **xlsx:** Inherit and preserve chart text formatting ([b292d16](https://github.com/ChristopherVR/ooxml/commit/b292d167b07c6cebaa9168f7a01ab113135d3857))
- **visio:** Preserve protected angles during flips ([48c9156](https://github.com/ChristopherVR/ooxml/commit/48c91568f14905441b152cefaea3b8486f99f86f))
- **xlsx:** Measure chart title and legend text in the browser ([e74cf9e](https://github.com/ChristopherVR/ooxml/commit/e74cf9e85fec51ee29fee5c2ad9f892c38a1c12f))
- **visio:** Retain guarded flip flags in native transforms ([ed7c987](https://github.com/ChristopherVR/ooxml/commit/ed7c98774e3e30292348c536b222e89c64bb408b))
- **xlsx:** Align chart title lines with measured font boxes ([8499a58](https://github.com/ChristopherVR/ooxml/commit/8499a584ea11a03cb25517f3d6dad3b16104b375))
- **ui:** Preserve feedback and history for blocked visio flips ([9feb751](https://github.com/ChristopherVR/ooxml/commit/9feb75113d0e0d3f01bf509bbe55d08ee46af7ce))
- **docx:** Remove incomplete field structure from clipboard ([1539bbc](https://github.com/ChristopherVR/ooxml/commit/1539bbc9a136a904aad1e3b73f59834ac195ccd8))
- **docx:** Preserve explicitly empty simple field caches ([a96867f](https://github.com/ChristopherVR/ooxml/commit/a96867f4d784d723812a5e2cf13a1eec8b343acf))
- **docx:** Preserve fields when deleting final graphemes ([8288b60](https://github.com/ChristopherVR/ooxml/commit/8288b605b3eb4358cc803ea05457946443174f6f))
- **docx:** Retain nested field result metadata ([768a5b5](https://github.com/ChristopherVR/ooxml/commit/768a5b580e814d7884211987fffadc3cf4353286))
- **docx:** Resolve logical alignment after direction overrides ([2cc0ae5](https://github.com/ChristopherVR/ooxml/commit/2cc0ae57f42a503b39c9b60cd5c63d8ddd901929))
- **docx:** Decode packed table look flags ([dbef394](https://github.com/ChristopherVR/ooxml/commit/dbef3940b521a5b6f37c40573e9b1528aba9d088))
- **docx:** Preserve field results across story paragraphs ([3b0cd9c](https://github.com/ChristopherVR/ooxml/commit/3b0cd9c2a526ea645cb7dfff56244b220c35549e))
- **docx:** Honor explicit off header and footer flags ([6fda7a3](https://github.com/ChristopherVR/ooxml/commit/6fda7a38d34343730821ebb8f81a4232f02fef44))
- **docx:** Project copied field result objects as literals ([dff5d16](https://github.com/ChristopherVR/ooxml/commit/dff5d168240dac773478dba6bc86dd3758c94a03))
- **docx:** Follow word numbering override restart rules ([a5f4daf](https://github.com/ChristopherVR/ooxml/commit/a5f4daf04f62b6ea81682c65ae9f975a766db3c8))
- **ci:** Exclude core test support from published declarations ([11666e6](https://github.com/ChristopherVR/ooxml/commit/11666e65e2f9dc9f02f31add14fea70c8e29a51c))
- **docx:** Preserve and resolve omitted numbering starts ([cde7d37](https://github.com/ChristopherVR/ooxml/commit/cde7d37fac28063ef233550feed3786bc1eadcd8))
- **ci:** Clear lint errors in chart and drawing checks ([f331e67](https://github.com/ChristopherVR/ooxml/commit/f331e670ca9c52a76e048648e55d0660d925f7c4))
- **pptx:** Preserve rectangular chart gradient paths ([27c6ddb](https://github.com/ChristopherVR/ooxml/commit/27c6ddbf0b123266a67c794b4bbaa0b7c0d062ef))

### Performance

- **pptx:** Share the engine and load metafile conversion lazily ([e00b1d7](https://github.com/ChristopherVR/ooxml/commit/e00b1d748b9adb79dac8579f67db24e75f0a7fe2))

### Testing

- **visio:** Measure rotated gradients in native raster frames ([e47981c](https://github.com/ChristopherVR/ooxml/commit/e47981c88dfb7d2c689c82e2c7c99d9232bc29c9))
- **visio:** Capture native line translation formulas ([79c274d](https://github.com/ChristopherVR/ooxml/commit/79c274d7ccd6c263a50ad464737739368de9ca62))
- **xlsx:** Measure native linear gradient paint profiles ([ccc4ceb](https://github.com/ChristopherVR/ooxml/commit/ccc4cebc6408a47f73e85ce95262ebe61973deda))
- **visio:** Verify endpoint edits across native drawing scales ([6f02467](https://github.com/ChristopherVR/ooxml/commit/6f02467063663351dbd73d84d125d9571e8dbf11))
- **visio:** Verify native endpoint gradient paint after saving ([be42b4d](https://github.com/ChristopherVR/ooxml/commit/be42b4d75e80266c63fa9e34ec46d42b7de50f43))
- **visio:** Verify native gradient paint after endpoint gestures ([c5d868d](https://github.com/ChristopherVR/ooxml/commit/c5d868dbcc432a2c219fd210a8d8c0355e17e039))
- **visio:** Record native endpoints after width overrides ([c699b54](https://github.com/ChristopherVR/ooxml/commit/c699b549a2fc4d7cb1fa57b257cf5efc1a919bb1))
- **visio:** Remove unused group rotation import ([906e7ef](https://github.com/ChristopherVR/ooxml/commit/906e7efa554bc964242ac441803a257c2ffb97e7))

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
