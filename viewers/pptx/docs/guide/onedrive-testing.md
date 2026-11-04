---
title: Testing OneDrive integration
description: Verify the shared Office OneDrive connection with test-local transports and authorized Microsoft account acceptance.
---

# Testing OneDrive integration

The OneDrive connection is shared by Word, Excel, PowerPoint and Visio. Its automated transports
belong only to test fixtures. There is no shipped simulation UI, public mock service or OneDrive
simulation demo to configure or deploy.

Use the [Office testing guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)
for commands, per-format checks and the real-account acceptance matrix. The
[integration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)
documents package and framework wiring, authentication, sharing and conflict recovery.

Automated regression tests can verify the Graph request sequence and editor save wiring without
Microsoft credentials. They cover snapshots and file identity, permission restrictions, revision
conditions, retained local edits, separate cloud copies and error handling. They cannot prove
Microsoft sign-in, consent, Conditional Access, deployed-origin CORS or native PowerPoint fidelity.

Microsoft Graph Explorer allows manual API checks using Microsoft's existing application
registration. It does not validate the application's own registration or sign-in flow. Use a
disposable deck and follow the testing guide; do not ship Graph Explorer's client ID or borrowed
credentials as an application authentication solution.

For actual editing through the suite, configure the host's Microsoft token provider before its
entry loads, then use **Open from OneDrive**. Real personal/work accounts and SharePoint tenant
policies remain required acceptance cases. No registration-free test is evidence of live
Microsoft coauthoring or guaranteed conflict-free concurrent editing.
