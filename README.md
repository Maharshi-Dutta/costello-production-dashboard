# Costello Production Dashboard

Live view of the production workbook on SharePoint, for Costello Windows.

- Reads the workbook through Microsoft Graph, including cell fill colours
  (gold `#FFE699` = ready to deliver, yellow `#FFFF00` = in fabrication).
- Writes status changes straight back into the sheet, cell by cell.
- Each person signs in with their own Costello account; they see exactly the
  jobs they can already open in Excel.

## Files
| file | purpose |
|---|---|
| `index.html` | page shell and styles |
| `graph.js` | Entra sign-in (MSAL) and all Graph read/write calls |
| `parser.js` | workbook → job model; runs in the browser and in Node |
| `app.js` | dashboard UI |
| `verify.js` | dev check: compares the parser against a reference extract |

Rows in the sheet move, so every write re-finds the job by its job number
immediately before writing. Nothing caches a row position.
