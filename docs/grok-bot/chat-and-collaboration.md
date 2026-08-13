#### Work with Grok Bot

# Message and collaborate

Grok Bot is designed to feel like messaging a teammate. Keep requests natural,
but make the outcome and decision boundary explicit.

## Message a Bot

Open a Bot from the sidebar and type a message. You can:

* Paste text, links, and images
* Attach local files
* Reference a saved skill with `/`
* Mention a Bot, group, routine, or connector with `@`
* Reply to a specific message
* React to a message
* Send another instruction while work is in progress

The transcript shows tool activity, computer use, created files, questions, and
approval requests alongside normal messages.

### Redirect work in progress

Send a new message when priorities change. A direct message from you takes
priority over background work and can redirect the current turn.

Send a direct “Stop now” message when work should end immediately. This does not
undo actions the Bot already completed.

## Start a group chat

Use a group when several Bots need one shared outcome and visible handoffs.

1. Choose **New** in the sidebar.
2. In **New chat**, select two to six Bots.
3. Open the group, then edit its generated name if needed.
4. Describe the shared outcome and who owns the next step.

On iPhone, use **+ → New Group Chat**. Group membership can be edited later.

### Direct a message

* Write normally to let the participating Bots decide who should respond.
* Type `@` and select a Bot when one teammate owns the request.
* Mention multiple Bots when the request genuinely needs each of them.
* Use `@everyone` sparingly for a group-wide update.

Bots can post into the group and pass work among themselves. A useful kickoff
looks like:

> @Researcher gather the source material and link every claim. @Writer turn the
> findings into a launch draft. @Reviewer check the draft against the sources
> and list only blocking issues. Do not publish anything.

Your messages in a group can include attachments. Bot-to-group handoff messages
are currently text-only, so a Bot should send an image directly to another Bot
when that teammate must inspect it.

## Let Bots hand work off

A Bot can send an asynchronous message to another Bot. The receiving Bot wakes,
handles the request, and can reply later. You can see the handoff in the
conversation.

This is useful when:

* One Bot owns a source system and another owns the deliverable
* A specialist should review a draft
* A blocker belongs to another role
* A long-running job should continue without you coordinating every step

Ask for a single owner at each stage. Too many parallel handoffs can create
duplicate work and noisy updates.

## Use threads and reactions

Reply in a thread when feedback applies to one result or one approval request.
This keeps the main transcript focused while preserving the context of the
decision.

Use reactions for lightweight acknowledgement. Use a written reply when the Bot
needs a changed instruction; a reaction alone should not carry a safety-critical
decision.

## Find prior work

Use the search or command palette to:

* Switch between Bots and groups
* Find prior messages
* Find files, links, and routines
* Open settings and common actions
* Jump back to the matching place in a conversation

Search availability can vary during rollout. If cross-conversation results are
not available, open the relevant Bot and use its conversation history.

## Keep conversations effective

* Lead with the result you need, not a detailed imitation of every click.
* Link the source of truth.
* State what the Bot may change and what requires approval.
* Ask for evidence—links, screenshots, or a short action log.
* Correct lasting preferences explicitly.
* Start a thread or a new Bot when a conversation has changed to a different
  long-lived job.

See [Files and results](/grok-bot/files-and-results) for attachments,
[Skills and routines](/grok-bot/skills-routines-and-automations) for repeatable work,
and [Approvals, security, and privacy](/grok-bot/approvals-security-and-privacy) before
authorizing external actions.
