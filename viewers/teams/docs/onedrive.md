# OneDrive links and Office documents

For a step-by-step setup (Microsoft Entra registration, sign-in and a complete example), start with
the [OneDrive quick start](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md).

OpenTeams can carry a OneDrive link to a DOCX, XLSX, PPTX or VSDX file. The integrated Office
Suite opens that file through the matching document editor and the shared
`ooxml-core/embedding/onedrive` connection. OpenTeams itself is a workspace, not a fifth Office
file editor and not Microsoft Teams.

The same distinction applies to `openteams-react-viewer`, `openteams-vue-viewer`,
`openteams-angular-viewer`, `openteams-svelte-viewer`, `openteams-solid-viewer`,
`openteams-vanilla-viewer` and the `ooxml-ui/teams` element:

| Shared item                 | Where edits are stored                                                                          |
| --------------------------- | ----------------------------------------------------------------------------------------------- |
| OneDrive link               | In the original OneDrive item only when the authorized host saves through its OneDrive session. |
| Local workspace file link   | In the local Office workspace; it is not a cross-device cloud link.                             |
| OpenTeams server attachment | In that server's storage; an uploaded copy is independent of the OneDrive original.             |

The suite's connected-file Share flow requests an explicit view/edit role and link scope.
Posting or copying an existing item URL does not itself grant another person access. Never create
an anonymous link as a fallback when a tenant rejects a restricted link. `openteams-server`
does not acquire Microsoft credentials, mirror drive folders or merge Office file changes.

Use the [Office OneDrive integration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#sharing-and-openteams)
for authentication, package references and permission-aware sharing. Its
[testing guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md) distinguishes
test-local transports from real account acceptance. No simulation UI or service is shipped.

Yjs chat/presence and document collaboration rooms are separate from Microsoft Office coauthoring.
Sending the same OneDrive URL to two people does not automatically synchronize their editors.
