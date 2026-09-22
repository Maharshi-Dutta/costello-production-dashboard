# Job alerts by email

## 9. Job alerts by email

**What.** The admin subscribes addresses to jobs; every third day at 08:00 an
email per address lists its jobs still in production, job number and
comments only.

**How.** Subscriptions live in `Dashboard Alerts` (Job|Email|Added by|When);
only the admin (client-side check against `Dashboard Config` key `admin`)
can add or remove; recipients must be on the admin's domain. The mailer is an
Office Script (`automation/alerts-digest.ts`, generated from
`automation/digest-core.js` by `build-script.js`) run by a Power Automate
flow the owner creates by hand from `automation/SETUP.md`, sending from the
admin's mailbox. No address or domain is in the repo; they come from
`Dashboard Config` at run time. Tests: `test_alerts.js`,
`automation/test_digest.js`.

## See also

- [[ui-extras]] — previous: the selection wheel and versions
- [[export]] — next: export to Excel / PDF
