# Quick Schedule Send

Outlook add-in for **Innovation Schoolwear** (UK, Europe/London). In message compose it puts a **Schedule send** group on the Message tab: **1h**, **2h**, **3h**, **4h**, **Tomorrow 8am**, and a gear **Edit times** task pane.

Clicking a time only **chooses** the delay. The clock starts when you press **Send**. An `OnMessageSend` handler (Smart Alerts, `PromptUser`) reads the choice and sets `item.delayDeliveryTime` to now + delay, then allows the send.

This repo is a static site (HTML/JS/CSS). There is no backend.

## What you get

| Path | Role |
| --- | --- |
| `src/` | Host this folder on any HTTPS static host |
| `manifest.xml` | Add-in only manifest (use this to sideload / deploy) |
| `manifest.json` | Unified Microsoft 365 manifest (optional; Mac still needs the XML) |
| `host.config.json` | Single host constant |
| `scripts/set-host.mjs` | Writes both manifests from that host |
| `src/assets/` | Icons 16 / 32 / 64 / 80 / 128 (plus gear and unified colour/outline) |

## Set the host and regenerate the manifests

Office add-ins must load from **HTTPS**. Put the contents of `src/` on a host, then:

```bash
node scripts/set-host.mjs https://your-static-host.example/quick-schedule-send
```

That updates `host.config.json` and rewrites `manifest.xml`, `manifest.json`, and the concatenated event file `src/launchevent.js`. Running the script with no argument regenerates from the host already in `host.config.json`.

The placeholder in a fresh clone is `https://addin.example.com`. Do not sideload until that URL actually serves `src/`.

## Host the static folder

Upload everything in `src/` so these URLs resolve:

- `…/taskpane.html`
- `…/commands.html`
- `…/launchevent.js`
- `…/assets/icon-16.png` (and 32/64/80/128)

Any HTTPS static host works (SharePoint document library, Azure Blob/Static Web Apps, GitHub Pages, Netlify, an IIS site). CORS is not required for Office to load your pages. The host must match the URLs in the generated manifest.

Local UI preview (not for Outlook itself — Outlook will not load `http://`):

```bash
npm start
```

Then open the task pane page on port **43173**.

## Sideload for one user (new Outlook and Outlook on the web)

Microsoft’s current wording (Learn + Microsoft Support, 2026): **the new Store does not install a custom XML file**.

In new Outlook, **More apps → Add apps** (or **Apps → Add apps**) opens the Store. That path is **not** how you sideload this add-in. Support notes that installing a custom add-in from an XML file is only available through the older **Add-Ins for Outlook** dialog; they are still working on “Add from file” in the new Store.

**Do this instead:**

1. In a browser, open [https://aka.ms/olksideload](https://aka.ms/olksideload). Outlook on the web opens, then the **Add-Ins for Outlook** dialog.
2. Select **My add-ins**.
3. Under **Custom Addins**, select **Add a custom add-in** → **Add from File**.
   - **Add from URL** has been removed. Download the manifest first if you only have a link.
4. Choose the generated `manifest.xml` and accept the prompts (**Install**).

The sideload applies to Outlook on the web and should appear in new Outlook on Windows (and other supported desktop clients) for that mailbox. Classic Outlook on Windows can take up to 24 hours because of caching.

To remove it, use the same dialog → **Custom Addins** → **…** → **Remove**.

## Org-wide deployment

1. Microsoft 365 admin centre → **Settings** → **Integrated apps**.
2. **Upload custom apps**.
3. App type **Office Add-in**, upload `manifest.xml` (or the unified `manifest.json` if your tenant uses that flow).
4. Assign users or groups and finish the wizard.

### OnMessageSend / event-based activation — is admin deployment required?

**For production use: yes, unless the add-in later has an unrestricted Microsoft Marketplace listing.**

Microsoft’s event-based activation rules:

- If users install the add-in themselves from Microsoft Marketplace / the Office Store, **event-based handlers do not auto-launch**, so `OnMessageSend` will not run and the delay will not be applied.
- **Admin deployment** via Integrated apps is the supported way to turn event-based activation on for the organisation.
- An **unrestricted** Marketplace listing is the documented exception (event-based activation works after a user install). This add-in is not listed there.
- A **restricted** Marketplace listing still needs admin deployment for events.

Sideloading (`aka.ms/olksideload` or `npm start` in a Yeoman project) is what Microsoft’s Smart Alerts walkthrough uses for **development testing**. Treat one-user sideload as a test path, not as how Innovation Schoolwear should run this in production.

Related admin control: if people send while Outlook started offline, web / new Outlook cannot see which add-ins are installed. Admins can set the Exchange Online mailbox policy **`OnSendAddinsEnabled`** so send-time add-ins still run in that situation. See [Handle OnMessageSend](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/onmessagesend-onappointmentsend-events).

This add-in uses send mode **`PromptUser`** (not SoftBlock). The handler always allows send after it sets the delay (or if no delay was chosen). If setting the delay fails, send is still allowed and Outlook may show the error text.

## How the delay works

1. Ribbon or task pane stores the chosen preset on **this draft** (`item.sessionData`, with custom properties as fallback).
2. A notification on the item says e.g. “Will send 2 hours after you press Send”. **Clear schedule** on that insight opens the task pane (the only notification action Office.js supports) and clears the choice. The task pane has the same clear control.
3. On **Send**, `onMessageSendHandler` reads the preset, computes **now + delay** (or tomorrow 08:00 in the user’s IANA timezone, defaulting to Europe/London), calls `item.delayDeliveryTime.setAsync`, then `event.completed({ allowEvent: true })`.
4. If no button was chosen, the handler does nothing and send continues.

`delayDeliveryTime` (Mailbox 1.13) is processed **on the server**. Outlook does **not** need to stay open for the message to go out. That is different from the native Outlook “Delay Delivery” checkbox, which is client-side and needs the client running.

In **new Outlook**, Outlook on the web, and Mac, a delay-delivery message sits in **Drafts** until the time; you can still edit it. In **classic** Outlook on Windows it does not appear in Outbox after Send — you only see it later in Sent Items.

## Ribbon labels vs your edited buttons

Office.js **cannot change ribbon button captions at runtime**. `Office.ribbon.requestUpdate` can enable or disable a control; it cannot change the label or icon. There is no `getLabel` callback in web add-ins (that exists only in older COM/VSTO ribbons).

So:

- Manifest labels stay **1h / 2h / 3h / 4h / Tomorrow 8am / Edit times**.
- **Edit my buttons** in the task pane is what you customise. Those labels appear on the task pane chips.
- Ribbon **clicks** do follow your saved list as far as the platform allows: slot 1 applies your first saved preset, slot 2 the second, and so on. If a slot has no user preset, it uses the manifest default. Keep the first five presets aligned with the ribbon captions if you do not want “1h” to fire a 90-minute delay.

## Requirements

- Mailbox **1.13** (`delayDeliveryTime`). Compose on Exchange Online.
- New Outlook on Windows and Outlook on the web (also classic Windows / new Mac UI for the same APIs).
- Permission: **ReadWriteItem**.

## Develop

```bash
npm test
npm run validate          # official office-addin-manifest validator (XML)
node scripts/set-host.mjs # regenerate manifests + launchevent.js
npm run icons             # rebuild PNGs if you change the generator
npm start                 # static preview on port 43173
```

Unit tests cover hour offsets, tomorrow 08:00 across midnight, and the UK BST↔GMT changes (last Sunday in March / October).

Event runtime: classic Outlook on Windows loads `launchevent.js` as a single file with no imports. `scripts/set-host.mjs` concatenates `src/lib/delay.js`, `src/lib/office-schedule.js`, and `src/launchevent-handler.js`. New Outlook and the web load `commands.html` instead.

## Limits (read before you rely on this)

1. **Ribbon captions are fixed** in the manifest. User edits apply to the task pane; ribbon slots map as described above.
2. **New Outlook may collapse** a multi-button add-in group to the group name plus a flyout. The group **Schedule send** is on the default Message compose tab (`TabDefault`) so it is not behind a separate “open add-in” click in the manifest. If the client still shows one overflow control, that is a platform ribbon change, not something this add-in can override.
3. **OnMessageSend in production needs admin deployment** (or an unrestricted Marketplace listing). User-installed Store copies will not run the send handler.
4. **Outlook does not need to stay open** for `delayDeliveryTime`. Keep the draft if you need to cancel in new Outlook / web (it stays in Drafts).
5. Notification **Clear schedule** can only open the task pane (`ShowTaskPane`). There is no ExecuteFunction action on insight notifications.
6. Unified JSON manifest is included but **Outlook on Mac does not support it** yet — deploy `manifest.xml` if anyone uses Mac.
