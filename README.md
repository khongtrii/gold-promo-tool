# Gold Promo desktop generator

Activate the project environment, then launch the desktop application:

```powershell
. .\.venv\Scripts\Activate.ps1
python -m src.desktop_app
```

Stage 1 requires the Gold Promo source workbook and one Master Data workbook
(default: `data/reference/master-data-file.xlsm`). It reads the `site-group`
sheet for Site Groups and `plan-goldpromo` for the promotion plan. It creates the operational templates and
`template_ag.xlsx`; select the returned AG report to create the final discount
templates. Stage 2 requires the Gold Promo source workbook and an Attribute
file; it uses the source workbook's `Network Configure` sheet. Every output uses the `.xls` extension and
the suffix `_ddmmyy_hhmmss` (for example, `template_ag_030826_154522.xls`).

When validation writes `NOTE ERR FROM MASTER DATA`, processing stops and a
copy of the selected source workbook is returned in the output folder. Its
metadata rows, row-7 header layout, data rows, and `Network Configure` sheet
are retained.

Before Stage 1 creates outputs, unmatched Gold Promo networks are presented
for Site Group review. Suggestions are calculated exclusively by comparing
`GOLD PROMO NETWORK EXPANDED` with each master Site Group's SITE list, never
from the source `SITE GROUP` value. A code is suggested only when both the
missing-store count and the extra-store count are at most five; otherwise the
code is left blank for the user to enter. A master Site Group code that has
already been assigned to an exact match in the processed source, or suggested
to an earlier network, is excluded from later network suggestions. This keeps
one `GOLD PROMO NETWORK EXPANDED` mapped to one distinct Site Group.

Stage 1 also provides an **Excluded SITE GROUP codes** list. Codes added with
the `+` button are removed from the master Site Group list before exact
matching and suggestions; select a code and use `−` to remove it from the
exclusion list.
