# Embedding the editors

This guide is for hosts that put an Office editor inside their own application: what to set, what
the editor emits, and which side owns each decision. It starts with permissions and grows with the
rest of the embedding contract (regions, capabilities, add-in ribbon tabs, host-owned storage).

## Read-only and roles

Two things decide whether a user can change a document, and they are deliberately separate.

- **The host's permission: `readOnly`.** Every full editor and viewer element and its framework
  bindings take `readOnly` (boolean, default `false`). `true` makes the document read-only; `false`
  is the same as not setting it. Word and Excel also emit `readonly-change` when it flips.
- **The collaboration role: `CollaborationOptions.user.role`** (`src/core/collab/identity.ts`, see
  `docs/collab-area.md`). A `viewer` role makes the editor read-only for as long as the session is
  joined, and edits from a `viewer` are never published.

There is no separate role prop. Both restrictions apply together: the editor is writable only when
the host allows it and the role allows it. In every product the role lock is tracked apart from the host
switch, and the effective value drives the UI and `readonly-change`. Turning `readOnly` off never lifts a `viewer` role, and
leaving the room restores whatever the host last set.

| Product    | Prop                                                               | Inverse alias (kept)                             |
| ---------- | ------------------------------------------------------------------ | ------------------------------------------------ |
| Word       | `readOnly`                                                         | none                                             |
| Excel      | `readOnly`                                                         | none                                             |
| Visio      | `readOnly`                                                         | `canEdit` (`canEdit: false` is `readOnly: true`) |
| PowerPoint | `readOnly` (`editable` in Svelte and vanilla, `canEdit` elsewhere) | the edit switch                                  |

If both spellings are set, **`readOnly: true` wins**: a deck or drawing is editable only when the
edit switch allows it and `readOnly` is not `true`. Setting `readOnly: false` does not grant edit
permission by itself. PowerPoint's edit switch defaults to off, so a PowerPoint host still passes
`canEdit` (or `editable`) to allow editing, and uses `readOnly` to take it away again, for example
per user.

Building blocks (the individual toolbar, canvas and pane elements) keep `canEdit`, read-only unless
it is `true`; they are not full editors and take no `readOnly`.

`customization.capabilities.edit` in PowerPoint is a third, independent restriction (a host policy
ceiling); like the role, it can only take editing away.

### Mirroring `readOnly` from the editor

Word and Excel emit `readonly-change` whenever read-only changes: the user toggled Editing or
Viewing, the host changed `readOnly`, or a collaboration role was joined or left. The detail is
`{ readOnly, locked, hostReadOnly }`:

- `readOnly` is the effective value (the host's switch or a `viewer` role).
- `locked` is true while a `viewer` role holds the editor read-only.
- `hostReadOnly` is the host's own switch.

A parent that keeps `readOnly` in its own state must store `hostReadOnly`, never `readOnly`;
otherwise a role lock would be written into the host's setting and leaving the room would not
restore it. The two-way channels do this for you: Angular's `[(readOnly)]` (`readOnlyChange`) and
Vue's `v-model:readOnly` (`update:readOnly`) carry the host switch only and stay quiet when only a
role lock changed. The callbacks (`onReadOnlyChange`, Vue's `readonly-change`) receive the
effective value first and the detail second. Word's `readonly-change` detail was a boolean before
and is now this object. Visio and PowerPoint have no read-only change event, so nothing there can
write a role lock back.
