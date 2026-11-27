# Using your own report format

The GST reconciliation report is built from an Excel template. If you do not
upload one, you get the standard layout (`Outputformat.xlsx`). You can upload
your own version instead: change the standard one, or design your own.

The reconciliation itself never changes with the template. Matching, statuses,
categories, ITC amounts and anomalies are worked out the same way for every
client. The template only decides how the results are laid out.

## What you can change

- **Columns**: add, remove, rename or reorder them on any sheet.
- **Sheets**: remove sheets, rename them, reorder them, or add your own.
- **Look**: fonts, colours, number formats, column widths, frozen panes and
  filters are kept exactly as you set them.
- **Calculated columns**: put an Excel formula in the sample row (see below)
  and it is copied down to every row.

## How a template is laid out

Each sheet that holds a list follows the same pattern as the standard
template:

1. Optional title and note rows at the top. They are kept. Placeholders in
   angle brackets such as `<n>`, `<Entity name>` or `<GSTIN>` are filled in.
2. **One header row**: the first row with two or more filled cells.
3. **One sample row** directly under the headers. It shows the formatting for
   every row of data and is replaced by the real data.
4. An optional totals row under the sample row, for example
   `=SUBTOTAL(9,O5:O5)`. Its ranges are extended to cover all the data.

Anything else under the sample row is cleared. If you build your template by
editing last month's report, its old rows are removed and never delivered,
but it is cleaner to keep just the headers and one sample row.

The **Summary** sheet (also recognised when named Dashboard or Overview) is
made of numbered sections such as `1. Data in scope` or `4. ITC To Be
Reclaimed`. You can remove a section, change its columns, or move it. A
section with a heading we do not recognise is left as you wrote it.

## How sheets and columns are recognised

A sheet is recognised by its name, and a column by its header, using the
accepted names in the tables below. Capital letters, spaces and punctuation
are ignored, `GSTR-2B` counts the same as `2B`, and `Purchase Register` or
`Books` count the same as `PR`. So `Taxable as per Books` and `PR Taxable`
are both fine.

Names are never guessed. If a header is not in the tables and is not
explained in a Column Guide, that column comes back **empty**, and the run
notes say which column and why. The report is still delivered.

Each sheet shows one kind of row (a "row set"). For example a sheet named
`Probable Match` lists only probable matches, and a sheet named `PR Only`
lists only documents found in the Purchase Register but not in GSTR-2B.

## Column Guide: explaining your own columns

If you use a header that is not in the tables, add a sheet called
**Column Guide** to your template with three columns:

| Sheet | Column | Meaning |
|---|---|---|
| All Recon | Supplier Ref | `pr_doc_ref` |
| All Recon | Books Value | invoice total in the Purchase Register |
| Vendor Issues |  | `not_matching` |

- **Meaning** is best written as a field name from the tables (exact), or in
  plain English. It must describe exactly one value.
- Leave **Column** empty to say what kind of rows a sheet holds (use a row set
  name such as `not_matching`).
- Calculations do not go in the Column Guide. Use a formula in the sample row.
- The Column Guide sheet is removed from the delivered report.
- Anything else written in the Column Guide (for example a request to change
  thresholds) is ignored and reported as an anomaly.

## Calculated columns

To add your own calculation, put the formula in the sample row under your
header, using the sample row's cells. Example: header `Tax %` with sample-row
formula `=IFERROR(S5/O5,0)`. The formula is copied to every row with the row
number adjusted. Only headers that are not a known field are treated this
way.

## What cannot be filled

- Values that are not in the Purchase Register or GSTR-2B files, such as cost
  centre, approver or payment status. These columns come back empty.
- New kinds of grouping, such as totals per supplier. A sheet like that is
  returned with its headers only, and the run notes say so.

## If nothing in the template can be filled

If no sheet and no column in your template can be recognised (for example the
wrong file was uploaded), the run stops with the message:

> The output template has no sheet or column that could be filled. Please
> check the template or its Column Guide.

## Fields and accepted names

Field names are shown in `code`. Any of the accepted names works as a column
header for that field on that kind of sheet.

### `documents`: one per All Recon row (a pair, a PR-only or a 2B-only document)

Sheet names: All Recon, Reconciliation, Recon, All Documents, Reconciliation Detail.
- `matched` (documents rows with Match Status Matched): sheet names Matched, Matched Documents; same fields.
- `probable` (documents rows with Match Status Probable Match): sheet names Probable Match, Probable, Probable Matches; same fields.
- `not_matching` (documents rows with Match Status Not Matching): sheet names Not Matching, Mismatch, Mismatches, Unmatched; same fields.
- `pr_only` (documents rows with Presence PR only): sheet names PR Only, In PR Only, Missing in 2B; same fields.
- `2b_only` (documents rows with Presence 2B only): sheet names 2B Only, In 2B Only, Missing in PR; same fields.
- `value_mismatch` (documents rows with category Value mismatch (same invoice no.)): sheet names Value Mismatch, Amount Mismatch; same fields.

| Field | Meaning | Accepted names |
|---|---|---|
| `recon_id` | Recon ID shared by the PR and 2B side of the row (R001...) | Recon ID; Recon No.; Ref ID |
| `match_status` | Matched / Probable Match / Not Matching | Match Status; Status; Recon Status; Result |
| `recon_category` | the one category from Step 3 | Recon Category; Category; Reconciliation Category; Mismatch Category |
| `presence` | Both / PR only / 2B only | Presence; Present In; Available In |
| `doc_type` | Invoice / Credit Note / Debit Note | Doc Type; Document Type; Type |
| `confidence` | High / Medium / Low for Probable rows; blank otherwise | Confidence; Confidence Level |
| `remarks` | match basis and what differs (Step 3 reason text) | Match Basis / Remarks; Remarks; Match Basis; Reason; Comments |
| `pr_id` | PR ID of the PR document (PR001...) | PR ID |
| `pr_period` | PR month of the PR document (MMM-YY) | PR Period; PR Month |
| `pr_doc_ref` | Doc Ref of the PR document | PR Doc Ref; Doc Ref; Document Reference; Voucher No.; PR Voucher No. |
| `pr_gstin` | supplier GSTIN in the PR | PR GSTIN; PR Supplier GSTIN; GSTIN as per PR |
| `pr_party` | supplier name in the PR | PR Party; PR Supplier; PR Supplier Name; PR Party Name; Party Name as per PR |
| `pr_invoice_no` | supplier invoice / note number in the PR | PR Invoice No.; PR Invoice Number; PR Inv No.; PR Bill No.; Invoice No. as per PR |
| `pr_invoice_date` | invoice date in the PR | PR Inv Date; PR Invoice Date; PR Bill Date; Invoice Date as per PR |
| `pr_taxable` | taxable value in the PR document | PR Taxable; PR Taxable Value; PR Taxable Amt; PR Taxable Amount; Taxable as per PR |
| `pr_igst` | IGST in the PR document | PR IGST; IGST as per PR |
| `pr_cgst` | CGST in the PR document | PR CGST; CGST as per PR |
| `pr_sgst` | SGST in the PR document | PR SGST; SGST as per PR |
| `pr_tax_total` | IGST + CGST + SGST in the PR document | PR Total Tax; PR Tax; PR Tax Amount; Total Tax as per PR |
| `pr_invoice_total` | invoice total in the PR document | PR Invoice Total; PR Invoice Value; PR Invoice Amount; PR Total Value; Invoice Value as per PR |
| `2b_id` | 2B ID of the 2B document (GB001...) | 2B ID |
| `2b_period` | 2B month of the 2B document (MMM-YY) | 2B Period; 2B Month; 2B Return Period |
| `2b_section` | GSTR-2B section (B2B / B2BA / B2B-CDNR) | 2B Section; Section |
| `2b_gstin` | supplier GSTIN in 2B | 2B GSTIN; 2B Supplier GSTIN; GSTIN as per 2B |
| `2b_supplier` | supplier name in 2B | 2B Supplier; 2B Supplier Name; 2B Party; Supplier Name as per 2B |
| `2b_invoice_no` | invoice / note number in 2B | 2B Invoice / Note No.; 2B Invoice No.; 2B Invoice Number; 2B Note No.; Invoice No. as per 2B |
| `2b_invoice_date` | document date in 2B | 2B Inv Date; 2B Invoice Date; 2B Doc Date; Invoice Date as per 2B |
| `2b_taxable` | taxable value in 2B | 2B Taxable; 2B Taxable Value; 2B Taxable Amt; 2B Taxable Amount; Taxable as per 2B |
| `2b_igst` | IGST in 2B | 2B IGST; IGST as per 2B |
| `2b_cgst` | CGST in 2B | 2B CGST; CGST as per 2B |
| `2b_sgst` | SGST in 2B | 2B SGST; SGST as per 2B |
| `2b_tax_total` | IGST + CGST + SGST in 2B | 2B Total Tax; 2B Tax; 2B Tax Amount; Total Tax as per 2B |
| `2b_doc_value` | document value in 2B | 2B Doc Value; 2B Invoice Value; 2B Invoice Total; 2B Document Value; Invoice Value as per 2B |
| `taxable_diff` | PR taxable minus 2B taxable | Taxable Diff (PR-2B); Taxable Diff; Taxable Difference; Difference in Taxable |
| `tax_diff` | PR total tax minus 2B total tax | Tax Diff (PR-2B); Tax Diff; Tax Difference; Difference in Tax |
| `invoice_total_diff` | PR invoice total minus 2B document value | Invoice Total Diff (PR-2B); Invoice Total Diff; Invoice Value Diff; Invoice Value Difference |
| `date_diff_days` | PR invoice date minus 2B document date, in days | Date Diff (days); Date Diff; Date Difference; Days Difference |
| `within_threshold` | Yes / No: taxable, total tax and invoice total differences all within config.thresholds | Within Threshold?; Within Threshold; Within Tolerance |
| `invoice_no_same` | Yes / No: PR and 2B invoice numbers equal ignoring case and spaces | Invoice No. Same?; Invoice No. Match; Same Invoice No. |

### `itc_reclaim`: one per ITC To Be Reclaimed line (Step 4)

Sheet names: ITC To Be Reclaimed, ITC Reclaimable, Eligible ITC, ITC Claimable.

| Field | Meaning | Accepted names |
|---|---|---|
| `recon_id` | Recon ID of the documents row | Recon ID; Ref ID |
| `match_status` | Match Status of the documents row | Match Status; Status |
| `2b_period` | 2B month | 2B Period; Period |
| `2b_gstin` | supplier GSTIN in 2B | Supplier GSTIN; GSTIN; 2B GSTIN |
| `2b_supplier` | supplier name in 2B | Supplier; Supplier Name; 2B Supplier |
| `2b_invoice_no` | invoice / note number in 2B | Invoice / Note No.; Invoice No.; Invoice Number; 2B Invoice No. |
| `2b_invoice_date` | document date in 2B | Doc Date; Invoice Date; 2B Inv Date |
| `doc_type` | Invoice / Credit Note / Debit Note | Doc Type; Document Type |
| `claim_condition` | Eligible - matched, or the conditional wording from Step 4 | Claim Condition; Condition; Eligibility |
| `pr_igst` | PR IGST | PR IGST |
| `pr_cgst` | PR CGST | PR CGST |
| `pr_sgst` | PR SGST | PR SGST |
| `2b_igst` | 2B IGST | 2B IGST |
| `2b_cgst` | 2B CGST | 2B CGST |
| `2b_sgst` | 2B SGST | 2B SGST |
| `itc_igst` | IGST to reclaim (credit notes negative) | ITC IGST; IGST ITC; Eligible IGST |
| `itc_cgst` | CGST to reclaim (credit notes negative) | ITC CGST; CGST ITC; Eligible CGST |
| `itc_sgst` | SGST to reclaim (credit notes negative) | ITC SGST; SGST ITC; Eligible SGST |
| `itc_total` | itc_igst + itc_cgst + itc_sgst | ITC Total; Total ITC; Eligible ITC |

### `itc_not_reclaim`: one per ITC Not To Be Reclaimed line (Step 4)

Sheet names: ITC Not To Be Reclaimed, ITC Not Reclaimable, Ineligible ITC, ITC Not Claimable.

| Field | Meaning | Accepted names |
|---|---|---|
| `recon_id` | Recon ID of the documents row | Recon ID; Ref ID |
| `match_status` | Match Status of the documents row | Match Status; Status |
| `pr_period` | PR month | PR Period; Period |
| `pr_gstin` | supplier GSTIN in the PR | Supplier GSTIN; GSTIN; PR GSTIN |
| `pr_party` | supplier name in the PR | Supplier; Supplier Name; PR Party |
| `pr_invoice_no` | invoice / note number in the PR | Invoice / Note No.; Invoice No.; Invoice Number; PR Invoice No. |
| `pr_invoice_date` | invoice date in the PR | Doc Date; Invoice Date; PR Inv Date |
| `doc_type` | Invoice / Credit Note / Debit Note | Doc Type; Document Type |
| `reason` | why the ITC is not to be reclaimed (the category, or PR tax in excess of 2B) | Reason ITC not to be reclaimed; Reason; Reason Not Claimable |
| `action_note` | what to do about it | Action / Note; Action; Note; Recommended Action |
| `pr_igst` | PR IGST | PR IGST |
| `pr_cgst` | PR CGST | PR CGST |
| `pr_sgst` | PR SGST | PR SGST |
| `2b_igst` | 2B IGST | 2B IGST |
| `2b_cgst` | 2B CGST | 2B CGST |
| `2b_sgst` | 2B SGST | 2B SGST |
| `not_claimable_igst` | IGST not to reclaim (credit notes negative) | Not-claimable IGST; IGST Not Claimable; Blocked IGST |
| `not_claimable_cgst` | CGST not to reclaim (credit notes negative) | Not-claimable CGST; CGST Not Claimable; Blocked CGST |
| `not_claimable_sgst` | SGST not to reclaim (credit notes negative) | Not-claimable SGST; SGST Not Claimable; Blocked SGST |
| `not_claimable_total` | the three not-claimable heads added | Not-claimable Total; Total Not Claimable; Blocked ITC |

### `anomalies`: one per anomaly (Step 5)

Sheet names: Anomalies, Anomalies & Observations, Observations, Exceptions.

| Field | Meaning | Accepted names |
|---|---|---|
| `anomaly_no` | running number 1, 2, 3... | #; No.; S.No.; Sr No |
| `severity` | High / Medium / Low / Info | Severity; Priority |
| `area` | what the anomaly is about (e.g. GSTIN / master data) | Area; Anomaly Type |
| `observation` | what was found, with counts and amounts | Observation; Description; Details; Anomaly |
| `reference` | Recon IDs concerned | Reference (Recon IDs); Reference; Recon IDs |
| `tax_impact` | tax amount involved (Rs) | Tax impact (Rs); Tax Impact; Impact |
| `recommended_action` | what to do | Recommended action; Action; Action Required |

### `journal_check`: one per period x tax head (Step 2.5 and Step 5)

Sheet names: ITC Journal Check, Journal Check, Journal vs 2B.

| Field | Meaning | Accepted names |
|---|---|---|
| `period` | month (MMM-YY) | Period; Month |
| `head` | IGST / CGST / SGST | Head; Tax Head |
| `journal_itc` | ITC booked via journal (INPUT lines) | Journal ITC (INPUT lines); Journal ITC; ITC as per Journal |
| `pr_invoice_tax` | PR invoice-level tax, invoices minus credit notes | PR invoice-level tax (invoices - credit notes); PR Invoice Tax; PR Tax |
| `2b_itc` | GSTR-2B ITC, invoices + B2BA minus credit notes | GSTR-2B ITC (invoices + B2BA - credit notes); 2B ITC; 2B Tax |
| `journal_minus_2b` | journal_itc minus 2b_itc | Journal - 2B; Journal vs 2B |
| `journal_minus_pr` | journal_itc minus pr_invoice_tax | Journal - PR; Journal vs PR |
| `pr_minus_2b` | pr_invoice_tax minus 2b_itc | PR - 2B; PR vs 2B |

### `pr_data`: one per PR document, rolled up from lines (Step 2)

Sheet names: PR Data, Purchase Register, PR.

| Field | Meaning | Accepted names |
|---|---|---|
| `pr_id` | PR ID (PR001...) | PR ID |
| `pr_period` | PR month (MMM-YY) | PR Period; Period; Month |
| `pr_doc_ref` | Doc Ref | Doc Ref; Document Reference; Voucher No. |
| `source_rows` | PR file row numbers rolled into this document | Source rows (file); Source Rows; Source |
| `pr_gstin` | supplier GSTIN | Supplier GSTIN; GSTIN; GSTIN of Supplier |
| `pr_party` | supplier name | Party Name; Supplier; Supplier Name; Party |
| `pr_invoice_no` | supplier invoice / note number | Supplier Invoice No.; Invoice No.; Invoice Number; Bill No. |
| `match_key` | invoice number after Pass 2 normalisation | Match Key; Normalised Invoice No. |
| `pr_invoice_date` | invoice date | Doc Date; Invoice Date; Bill Date |
| `txn_date` | transaction / posting date | Txn Date; Transaction Date; Posting Date |
| `doc_type` | Invoice / Credit Note / Debit Note | Doc Type; Document Type |
| `tariff_account` | tariff or account of the lines | Tariff / Account; Account; Tariff; Ledger |
| `gst_group` | GST group or rate | GST Group; Tax Group; GST Rate |
| `line_count` | number of PR lines rolled up | Lines; Line Count; No. of Lines |
| `pr_base` | base value | Base Value; Base; Base Amount |
| `pr_taxable` | taxable value | Taxable Value; Taxable; Taxable Amount; Taxable Amt |
| `pr_igst` | IGST | IGST |
| `pr_cgst` | CGST | CGST |
| `pr_sgst` | SGST | SGST |
| `pr_tax_total` | IGST + CGST + SGST | Total Tax; Tax; Tax Amount |
| `pr_invoice_total` | invoice total | Invoice Amount; Invoice Total; Invoice Value |
| `rcm_ref` | Yes / No: reverse-charge reference in the invoice number (Step 2.3) | RCM ref in invoice no.; RCM Ref |
| `recon_id` | Recon ID of the documents row | Recon ID; Ref ID |
| `match_status` | Match Status of the documents row | Recon Status; Match Status; Status |

### `2b_data`: one per GSTR-2B document (B2B, B2BA, B2B-CDNR)

Sheet names: 2B Data, GSTR-2B, 2B.

| Field | Meaning | Accepted names |
|---|---|---|
| `2b_id` | 2B ID (GB001...) | 2B ID |
| `2b_period` | 2B month (MMM-YY) | 2B Period; Period; Month; Return Period |
| `2b_section` | section | Section; 2B Section |
| `2b_gstin` | supplier GSTIN | Supplier GSTIN; GSTIN; GSTIN of Supplier |
| `2b_supplier` | supplier trade / legal name | Supplier Name; Supplier; Trade Name; Legal Name |
| `2b_invoice_no` | invoice / note number (B2BA: revised) | Invoice / Note No.; Invoice No.; Invoice Number; Note No. |
| `original_no` | original number for B2BA | Original No. (B2BA); Original No.; Original Invoice No. |
| `2b_invoice_date` | document date | Doc Date; Invoice Date; Note Date |
| `doc_type` | Regular / Credit Note / Debit Note as in 2B | Doc Type; Document Type; Invoice Type |
| `reverse_charge` | Yes / No | Reverse Charge; RCM |
| `2b_taxable` | taxable value | Taxable Value; Taxable; Taxable Amount; Taxable Amt |
| `2b_igst` | IGST | IGST |
| `2b_cgst` | CGST | CGST |
| `2b_sgst` | SGST | SGST |
| `2b_cess` | cess | Cess |
| `2b_tax_total` | IGST + CGST + SGST + cess | Total Tax; Tax; Tax Amount |
| `2b_doc_value` | document value | Doc Value; Invoice Value; Invoice Total; Document Value |
| `gstr1_period` | supplier GSTR-1 period | GSTR-1 Period; Filing Period |
| `gstr1_filing_date` | supplier GSTR-1 filing date | GSTR-1 Filing Date; Filing Date |
| `itc_availability` | ITC availability as in 2B | ITC Availability; ITC Available |
| `recon_id` | Recon ID of the documents row | Recon ID; Ref ID |
| `match_status` | Match Status of the documents row | Recon Status; Match Status; Status |

### `summary_scope`: one per source (Purchase Register, GSTR-2B) per month; "2B not provided" for a missing month

Sheet names: Data in Scope.
Summary section headings: Data in scope.

| Field | Meaning | Accepted names |
|---|---|---|
| `source` | Purchase Register or GSTR-2B | Source |
| `period` | month (MMM-YY) | Period; Month |
| `invoice_count` | number of invoices | Invoices (#); Invoices; Invoice Count; No. of Invoices |
| `credit_note_count` | number of credit notes | Credit notes (#); Credit Notes; Credit Note Count; No. of Credit Notes |
| `taxable` | taxable value, net of credit notes | Taxable; Taxable Value |
| `igst` | IGST, net of credit notes | IGST |
| `cgst` | CGST, net of credit notes | CGST |
| `sgst` | SGST, net of credit notes | SGST |
| `tax_total` | IGST + CGST + SGST | Total tax; Tax; Tax Amount |

### `summary_status`: fixed rows Matched/Both, Probable Match/Both, Not Matching/Both, Not Matching/PR only, Not Matching/2B only, then Total and the control check

Sheet names: Reconciliation Result, Recon Summary.
Summary section headings: Reconciliation result.

| Field | Meaning | Accepted names |
|---|---|---|
| `match_status` | status of the row | Status; Match Status |
| `presence` | presence of the row | Presence |
| `document_count` | number of documents rows | Documents (#); Documents; Count; No. of Documents |
| `share_of_rows` | document_count divided by the total | % of all rows; % of Total; Share |
| `pr_taxable` | PR taxable | PR Taxable |
| `pr_tax_total` | PR total tax | PR Total Tax; PR Tax |
| `2b_taxable` | 2B taxable | 2B Taxable |
| `2b_tax_total` | 2B total tax | 2B Total Tax; 2B Tax |
| `tax_diff` | PR total tax minus 2B total tax | Tax diff (PR-2B); Tax Diff; Tax Difference |
| `control_check` | total rows minus All Recon rows; must be 0 | Check; Control Check |

### `summary_category`: one per recon category, then a Total row

Sheet names: Category Summary, Category Breakdown.
Summary section headings: Breakdown by reconciliation category, Breakdown by category, Category breakdown.

| Field | Meaning | Accepted names |
|---|---|---|
| `recon_category` | category | Category; Recon Category |
| `match_status` | status of that category | Status; Match Status |
| `document_count` | number of documents rows | Documents (#); Documents; Count; No. of Documents |
| `pr_tax_total` | PR total tax | PR Total Tax; PR Tax |
| `2b_tax_total` | 2B total tax | 2B Total Tax; 2B Tax |

### `summary_itc_reclaim`: one per claim condition, then Total ITC To Be Reclaimed

Sheet names: ITC Reclaim Summary.
Summary section headings: ITC To Be Reclaimed.

| Field | Meaning | Accepted names |
|---|---|---|
| `claim_condition` | claim condition | Claim condition; Condition |
| `document_count` | number of lines | Documents (#); Documents; Count |
| `itc_igst` | IGST to reclaim | IGST; ITC IGST |
| `itc_cgst` | CGST to reclaim | CGST; ITC CGST |
| `itc_sgst` | SGST to reclaim | SGST; ITC SGST |
| `itc_total` | total ITC to reclaim | Total ITC; ITC Total; Total |

### `summary_itc_not_reclaim`: one per reason, then Total ITC Not To Be Reclaimed and "of which pending (2B not provided)"

Sheet names: ITC Not Reclaim Summary.
Summary section headings: ITC Not To Be Reclaimed.

| Field | Meaning | Accepted names |
|---|---|---|
| `reason` | reason | Reason |
| `document_count` | number of lines | Documents (#); Documents; Count |
| `not_claimable_igst` | IGST not to reclaim | IGST; Not-claimable IGST |
| `not_claimable_cgst` | CGST not to reclaim | CGST; Not-claimable CGST |
| `not_claimable_sgst` | SGST not to reclaim | SGST; Not-claimable SGST |
| `not_claimable_total` | total ITC not to reclaim | Total ITC; Not-claimable Total; Total |

### `summary_anomalies`: fixed rows High, Medium, Low, Info

Sheet names: Anomaly Summary.
Summary section headings: Anomalies logged, Anomalies.

| Field | Meaning | Accepted names |
|---|---|---|
| `severity` | High / Medium / Low / Info | Severity |
| `anomaly_count` | number of anomalies of that severity | Count; Anomalies (#); Number |

### `header_values`: single values for titles and placeholders

Used to fill placeholders in titles and notes, not as columns.

| Field | Meaning | Accepted names |
|---|---|---|
| `entity_name` | legal name from the GSTR-2B Read me sheet | Entity name; Entity; Legal Name |
| `entity_gstin` | entity GSTIN from the GSTR-2B Read me sheet | GSTIN; Entity GSTIN |
| `pr_range` | first to last PR month (MMM-YYYY to MMM-YYYY) | PR range |
| `2b_range` | first to last 2B month (MMM-YYYY to MMM-YYYY) | 2B range |
| `threshold_taxable` | config.thresholds.taxable_value | Threshold taxable |
| `threshold_tax` | config.thresholds.tax_amount | Threshold tax |
| `threshold_invoice_total` | config.thresholds.invoice_total | Threshold invoice total |
| `row_count` | number of data rows on the sheet whose title holds <n> | n; Row count |
