# Migrating to 2.0

## `ooxml-ui/pptx`: Tab in an inline text editor

`mapInlineEditorChord` now returns `tab` for Tab (and `suppress` for Shift+Tab), so Tab types a
tab character in a text box instead of moving focus away, as in PowerPoint. An editor that closed
on every result other than `suppress` would now close on Tab.

Before:

```ts
const chord = mapInlineEditorChord(event);
if (chord) {
	event.preventDefault();
	if (chord !== 'suppress') close();
}
```

After:

```ts
import { insertInlineEditorText, mapInlineEditorChord } from 'ooxml-ui/pptx';

const chord = mapInlineEditorChord(event);
if (chord) {
	event.preventDefault();
	if (chord === 'tab') insertInlineEditorText(document, '\t');
	else if (chord === 'commit') close();
}
```
