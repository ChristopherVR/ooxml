# Migrating to OOXML 2.0

Each section names the packages a breaking change affects and shows the code before and after.

## PowerPoint: pictures and media in a shared room

Packages: `ooxml-core/pptx`, `ooxml-ui/pptx`, PowerPoint viewers.

A shared presentation now carries its pictures and media in the room (`pptx:media`). The viewers
configure this themselves. Code that writes PowerPoint slides into a room with
`reconcileSlidesInYDoc` or `writeSlidesToYDoc`, and whose slides name embedded audio or video only
by part, tells the room where its page plays each part from, so a joiner without the file gets the
bytes.

Before:

```ts
import { readSlidesFromYDoc, reconcileSlidesInYDoc } from 'ooxml-ui/pptx';

reconcileSlidesInYDoc(slides, doc, factories);
const shared = readSlidesFromYDoc(doc);
```

After:

```ts
import {
	configureCollaborationMedia,
	readSlidesFromYDoc,
	reconcileSlidesInYDoc,
} from 'ooxml-ui/pptx';

// Once per room, before its slides are read or written.
configureCollaborationMedia(doc, { localMediaUrl: (path) => mediaDataUrls.get(path) });
reconcileSlidesInYDoc(slides, doc, factories);
// A picture is now a URL of this page, or an `ooxml-media:` token while its bytes are on the way.
const shared = readSlidesFromYDoc(doc);
```

A room written by 2.0 names its pictures by media token, so a viewer from before 2.0 in the same
room shows them as broken: upgrade every participant of a room together.
