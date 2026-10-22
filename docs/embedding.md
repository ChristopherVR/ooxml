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
