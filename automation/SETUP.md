# Job alert emails - setting up the flow

This is the click-by-click for the administrator. It has to be done **once**.
Nothing here is done by the dashboard: the dashboard only manages who is
subscribed to which job (the `Dashboard Alerts` sheet). The mail itself is sent
by a Power Automate flow that runs an Office Script over the workbook.

Placeholders below are written like `<this>` - substitute your own values. No
address, name, domain or link is stored in this repository.

- **Who sends the mail:** whoever creates the flow. It goes out from that
  mailbox, so create it as the administrator.
- **When:** every 3rd day at 08:00, Dublin time.
- **What each person gets:** one mail listing every job they are subscribed to
  that is still in the *In production* section, with that job's comments.
  Finished jobs simply stop appearing - there is no "finished" notice.

---

## Before you start

Two rows must exist on the **`Dashboard Config`** sheet (`Key` in column A,
`Value` in column B). Create the sheet by hand if it is not there:

| Key            | Value                                    |
| -------------- | ---------------------------------------- |
| `admin`        | `<the administrator's sign-in address>`  |
| `dashboardUrl` | `<the address the dashboard is served from>` |

The script reads these every run:

- `admin` decides the **only domain that may be emailed** - the part after the
  `@`, compared in lower case. Any subscription outside that domain is silently
  skipped, the same rule the dashboard enforces when the address is added.
- `dashboardUrl` is the link at the bottom of the mail. Only `http://` and
  `https://` links are used; anything else is ignored. If the row is missing,
  the mail simply has no link.

Also check that `Dashboard Alerts` exists and has at least one row. It is
created by the dashboard the first time the administrator adds a subscription.

---

## Part 1 - add the Office Script

1. Open the workbook in **Excel on the web** (in the browser, not the desktop
   app - Office Scripts only exist on the web).
2. Ribbon → **Automate** tab. If you cannot see it, Office Scripts is not
   licensed for your account and the rest of this cannot be done.
3. Click **New script**. The Code Editor opens on the right with a small
   starter script in it.
4. Select everything in the editor (`Ctrl`+`A`) and delete it.
5. Open `automation/alerts-digest.ts` from this repository, copy the **whole**
   file, and paste it into the editor.
6. Click the script's name at the top of the editor pane (it will say something
   like *Script 1*) and rename it to exactly:

   ```
   Dashboard alerts digest
   ```

7. Click **Save script**. There should be no red squiggles. If there are, the
   file was pasted incompletely - clear the editor and paste again.
8. Optional but worth doing: press **Run** once. It changes nothing in the
   workbook (the script only reads). The output pane at the bottom shows the
   JSON it produced. `[]` means "nobody has anything to be told about right
   now", which is a perfectly good result.

> The script never writes to the workbook and never sends anything. It reads
> four sheets - `Dashboard Config`, `Dashboard Alerts`, `Production`,
> `Dashboard Log` - and returns the mails as JSON text. The flow does the rest.

**Do not edit the script in Excel.** It is generated. If it ever needs a
change, change `automation/digest-core.js` in the repository, run
`node automation/build-script.js`, and paste the new file over the old one.

---

## Part 2 - build the flow

Go to **make.powerautomate.com** and sign in as the administrator.

### 2.1 Create it

1. **Create** → **Scheduled cloud flow**.
2. Flow name: `Dashboard alerts digest`.
3. Starting: today's date, **08:00**.
4. Repeat every **3** **Day**.
5. **Create**.

### 2.2 Fix the time zone

Open the **Recurrence** trigger → **Show advanced options**:

- **Time zone**: `(UTC+00:00) Dublin, Edinburgh, Lisbon, London`
- **At these hours**: `8`
- **At these minutes**: `0`

Setting the time zone matters: without it the schedule drifts by an hour when
the clocks change.

### 2.3 Run the script

**New step** → search for **Excel Online (Business)** → action **Run script**.

| Field            | What to choose                                          |
| ---------------- | ------------------------------------------------------- |
| Location         | **SharePoint Site** (not OneDrive)                       |
| Document Library | the library the workbook lives in                        |
| File             | browse to the production workbook (the file picker; do not type a path) |
| Script           | **Dashboard alerts digest**                              |

Once **Script** is chosen the action shows no further inputs - this script takes
no parameters.

### 2.4 Parse the result

**New step** → **Data Operation** → **Parse JSON**.

- **Content**: the dynamic content **result** from the *Run script* step. (In
  the expression editor it is `body('Run_script')?['result']`.)
- **Schema**: paste exactly this:

```json
{
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "email": {
                "type": "string"
            },
            "subject": {
                "type": "string"
            },
            "html": {
                "type": "string"
            }
        },
        "required": [
            "email",
            "subject",
            "html"
        ]
    }
}
```

### 2.5 Send one mail per address

**New step** → **Control** → **Apply to each**.

- **Select an output from previous steps**: the **Body** of the *Parse JSON*
  step.

Inside the loop: **Add an action** → **Office 365 Outlook** → **Send an email
(V2)**.

| Field    | Value                                    |
| -------- | ---------------------------------------- |
| To       | dynamic content **email**                |
| Subject  | dynamic content **subject**              |
| Body     | dynamic content **html**                 |

Then **Show advanced options** on the same action and set:

- **Is HTML**: **Yes**

(If *Is HTML* is left at *No*, recipients get a page of raw HTML tags.)

Leave *From (Send as)* empty - the mail goes from the account that owns the
flow.

### 2.6 Save

Click **Save** at the top. Power Automate refuses to save if a required field
is empty; fill it and save again.

---

## Part 3 - test it

1. Top right of the flow designer: **Test** → **Manually** → **Test**.
2. The run appears step by step. Green ticks all the way down means it worked.
3. Open the **Run script** step to see the JSON it returned, and the
   **Apply to each** step to see how many mails went out. If the array is empty
   nothing is sent - that is not a failure.
4. Check the administrator's **Sent Items** for the mails.

Sensible first test: subscribe your own address to one job that is currently in
the *In production* section, run the test, then remove the subscription again.

---

## Part 4 - living with it

**Run history.** Power Automate → **My flows** → *Dashboard alerts digest*. The
**28-day run history** on the flow's page lists every run with its status; click
a run to see each step's inputs and outputs, including the exact JSON the script
returned.

**Pause it.** On the same page, **Turn off** in the toolbar. The flow stops
scheduling immediately and keeps its history and definition; **Turn on** starts
it again from the next scheduled slot. Turning it off is the right move for a
shutdown week or if something starts mailing wrongly - it is instant and
reversible. Deleting the flow is not.

**Change who gets what.** Never here - in the dashboard. The administrator adds
and removes addresses on a job, and the next run picks the change up.

**Change the wording or the schedule.** Wording lives in the script (see the
note at the end of Part 1). The schedule lives in the Recurrence trigger.

---

## If something goes wrong

| Symptom | Cause | Fix |
| --- | --- | --- |
| No **Automate** tab in Excel | Office Scripts not licensed / not enabled | Ask the tenant administrator; nothing in this guide will work without it |
| Script shows red squiggles after pasting | partial paste | Clear the editor and paste the whole file again |
| Run script step fails with "file not found" | the workbook was moved or renamed | Reopen the step and pick the file again with the browser |
| Run script returns `[]` every time | no `admin` row on `Dashboard Config`, or no live subscriptions | Check the config sheet, then check `Dashboard Alerts` |
| Mail arrives full of `<p>` tags | **Is HTML** left at *No* | Set it to *Yes* on the Send an email (V2) action |
| Somebody is not getting mail | their jobs are all finished, or their address is outside the admin's domain | Check the job is still in *In production*, and that the address matches the `admin` row's domain |
| Mails arrive an hour early or late | the Recurrence time zone was not set | Set it to Dublin (2.2) |

---

## What is in the repository

| File | What it is |
| --- | --- |
| `automation/digest-core.js` | the logic: sections, comments, the digest. Plain JS, no Excel, no network |
| `automation/alerts-digest.ts` | **generated** - the Office Script you paste into Excel |
| `automation/build-script.js` | regenerates the `.ts` from the core: `node automation/build-script.js` |
| `automation/test_digest.js` | offline test of both: `node automation/test_digest.js` |
