# Saving lab records to a shared Google Sheet

The lab pages always save to the browser they run in. This document is only for the
optional second step: also posting a row to a Google Sheet, so both computers' sessions
end up in one place and later weeks can read the accumulated data.

Skip all of this if the **⬇ Download JSON** button is enough.

## What gets stored

One row per player per session, in tidy (long) format — one observation per row, which is
what pandas and every plotting library want. Fifteen columns:

`savedAt · lab · player · trials · flagged · falseStarts · median · mean · mid50Width ·
over280 · predTypicalMs · predOneNumber · claim · evidence · limitation`

## Setup (about 10 minutes, once)

1. Create a spreadsheet at <https://sheets.new> and name it something like
   `data-literacy-lab records`.
2. In that sheet: **Extensions → Apps Script**. Delete the placeholder code and paste the
   script below.
3. Change `SECRET` to any phrase you like.
4. **Deploy → New deployment → ⚙ → Web app.** Set *Execute as* to **Me**, and
   *Who has access* to **Anyone**. Deploy, then authorize when Google asks.
5. Copy the deployment URL — it ends in `/exec`.
6. In `week1/index.html`, find the `CONFIG` block near the top of the `<script>` and fill
   in both fields:

   ```js
   const CONFIG = {
     sheetUrl: "https://script.google.com/macros/s/AKfy.../exec",
     secret:   "the same phrase you set above",
   };
   ```

7. Commit and push. Press **Save to our sheet** at the end of a session to test it.

Whenever you edit the Apps Script afterwards, you must **Deploy → Manage deployments →
edit → New version**, or the old code keeps running.

## The script

```javascript
const SECRET = "change-me";   // must match CONFIG.secret in the lab page
const TAB    = "records";

function doPost(e) {
  const out = ContentService.createTextOutput()
                            .setMimeType(ContentService.MimeType.JSON);
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.secret !== SECRET) {
      return out.setContent(JSON.stringify({ok: false, error: "bad secret"}));
    }
    const r  = payload.record || {};
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);

    if (sh.getLastRow() === 0) {
      sh.appendRow(["savedAt", "lab", "player", "trials", "flagged", "falseStarts",
                    "median", "mean", "mid50Width", "over280",
                    "predTypicalMs", "predOneNumber",
                    "claim", "evidence", "limitation"]);
    }

    const pred = r.prediction || {};
    const cel  = r.cel || {};
    (r.players || []).forEach(function (p) {
      sh.appendRow([
        r.savedAt, r.lab, p.name,
        (p.trials  || []).join(" "),
        (p.flagged || []).join(" "),
        p.falseStarts, p.median, p.mean, p.mid50Width, p.over280,
        pred.typicalMs, pred.oneNumberEnough,
        cel.claim, cel.evidence, cel.limitation
      ]);
    });

    return out.setContent(JSON.stringify({ok: true, rows: (r.players || []).length}));
  } catch (err) {
    return out.setContent(JSON.stringify({ok: false, error: String(err)}));
  }
}
```

## Reading the data back in a notebook

**File → Share → Publish to web → comma-separated values**, then in Week 2 and later:

```python
import pandas as pd
df = pd.read_csv(PUBLISHED_CSV_URL)
```

That closes the loop the whole program is built on: the data Bonjun generates in week *n*
becomes the dataset he analyses in week *n+k*.

## Honest note on the secret

The page is public, so anyone who views its source can read `CONFIG.secret`. It stops
drive-by bots from appending junk rows; it is not real security. Keep the sheet to
reaction times and the sentences you write in it, and nothing that would matter if a
stranger read it.
