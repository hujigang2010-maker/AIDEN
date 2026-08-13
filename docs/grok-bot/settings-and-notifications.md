#### Manage and protect

# Settings and notifications

Use Grok Bot settings for app-wide and desktop-local behavior, and conversation
details for one Bot's profile and notifications.

## Open Grok Bot settings

Open the account menu and choose **Settings**, or press `Cmd/Ctrl+,`.

The **Grok Bot settings** dialog contains sections based on your account and
rollout. Some options described below may not appear.

## General

### Account

Sign in or out of the Cursor account used by Grok Bot. The account menu also
shows **About**, the installed Grok Bot version, and a link to the iOS app.

### Appearance

Choose **Follow System**, **Light**, or **Dark**.

### Agent

Configure shared and local Bot behavior:

* **Default Model**, when model selection is available
* **Timezone**, which routines use for schedules
* **Execution on Local Computer**
* **Auto-review** and personal approval rules

See [Approvals, security, and privacy](/grok-bot/approvals-security-and-privacy) before
allowing commands on your local computer.

**Execution on Local Computer** applies to the desktop in front of you.
Auto-review rules are stored on the current desktop and synced to its Grok Bot
computer; do not assume another desktop installation has the same rules.

## Plugins

Use **Marketplace** to discover connectors and packaged skills. Use **Yours** to
review installed plugins and private skills.

An installed connector may still need browser authentication. Connector tools
can be enabled or disabled individually. Team-provided plugins may be required
or restricted by an administrator.

See [Use the computer and apps](/grok-bot/computer-and-apps) for the connection flow.

## Usage and billing

When **Usage & Billing** is available, it shows weekly included usage and
on-demand usage for eligible non-enterprise accounts.

The account menu may also show **Weekly usage**. If neither surface appears,
review usage from the Cursor account page or contact the organization
administrator.

## Team Setup

Team members may see **Team Setup**. Administrators can provide managed setup
that runs on assigned Grok Bot computers; members can review or reinstall the
current setup.

Do not place secret values directly in managed setup instructions.

## Beta and updates

Use **Beta** for:

* **Check for Updates** or **Restart to Update** for the Grok Bot app
* **Update Agent Computer** to rebuild the shared computer while preserving
  durable state
* **Reset Agent Computer** as a last-resort recovery that may lose recent
  unsynced work
* Security-key or egress-routing options when available

The app update and Agent Computer update are separate. See
[Troubleshooting](/grok-bot/troubleshooting) for the safest recovery order.

## Edit one Bot

Open **View conversation details**, then **Agent settings**, to edit that Bot's:

* Name, title, and description
* Avatar
* **Notifications** preference

These settings belong to one Bot. **Execution on Local Computer** and
Auto-review settings are shared across Bots using the current setup, but are not
an account-synchronized policy across every device.

## Understand attention states

The Bot list distinguishes:

* **Needs attention** for a question, approval, or handoff
* **Unread activity** for a new result
* Working or typing status

Opening a conversation marks its current activity as read. Use the Bot menu to
mark a conversation read or unread manually.

## Control notifications

Turn on **Notifications** in a Bot's settings to receive an operating-system or
mobile notification when that Bot finishes or needs input. Group chats do not
have the same per-Bot notification switch.

Notifications are normally suppressed while Grok Bot is focused. The sidebar
and dock badge still show unread activity.

The iPhone app also asks for notification permission during first run. Both
device permission and the Bot's notification setting must allow the
notification. Mobile push delivery is rolling out and may not yet be enabled
for every account.

## Handle in-app errors

Errors appear above the composer in **Notifications**. You can dismiss one
notice or clear the list. Some notices include **Copy request ID** for support;
copy and share the complete ID.

Clearing a notice removes the notification, not the underlying external action
or Bot history.
