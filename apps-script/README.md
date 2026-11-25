# RSVP app (Google Apps Script)

The RSVP form, the couple's dashboard and the server code. These files run in Google Apps Script
(project "Wedding RSVP", bound to the "Wedding RSVPs" Google Sheet), not on GitHub Pages.
This folder is a copy for safekeeping: editing a file here changes nothing until it is pasted into Apps Script.

| File | What it is |
| --- | --- |
| `Code.gs` | Server code: saves replies to the sheet, phone lookup, dashboard data, alert and daily emails |
| `Index.html` | The guest RSVP form |
| `Dashboard.html` | The couple's password-protected dashboard |

## Settings (Project Settings → Script Properties, not in this code)
| Property | Value |
| --- | --- |
| `DASHBOARD_PASSWORD` | the dashboard password |
| `RSVP_OPEN` | `yes` (set `no` to close the form) |
| `SITE_URL` | `https://shahzadsorathia.github.io/VS-Wedding/` |
| `ALERT_EMAILS` | who gets the RSVP emails, comma separated |

No password, email address or guest data is stored in these files.
