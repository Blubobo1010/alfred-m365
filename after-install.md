alfred-m365 is installed (Outlook tools, toolset `outlook`). It acts on ONE mailbox and writes
nothing until you enable writes.

Next steps:
1. Set the mailbox (and the owner's name for confirmation prompts):
   `alfred config set plugins.entries.alfred-m365.settings.mailbox luca.betti@example.com`
   `alfred config set plugins.entries.alfred-m365.settings.owner_name Luca`
2. `alfred outlook test --other <another address of the company>` — logs in, reads the latest mail
   header and today's calendar, and checks that the other mailbox is DENIED (the Exchange RBAC
   scope must contain only the configured mailboxes).
3. Drafts and calendar: `alfred config set plugins.entries.alfred-m365.settings.write_enabled true`.
   Drafts land in Outlook's Drafts folder and are never sent. Creating or moving an appointment is
   shown to you first (prompt in the CLI, buttons in Telegram); with attendees, Outlook sends them
   the invitation in the mailbox owner's name.
4. Restart the gateway if it is running (`alfred gateway restart`).

The app has Mail.ReadWrite and Calendars.ReadWrite on that mailbox: Microsoft has no narrower
permission. This plugin itself never sends, moves or deletes mail and never deletes events.
