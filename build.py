import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from collections import defaultdict
import csv
import glob
import os

# All source CSVs (roster/seminar/scholarship snapshots, staged Gmail receipt threads) live here rather than in
# /tmp, since /tmp isn't guaranteed to survive between the daily scheduled refreshes (Aug 28, 2026 -- set up for
# Michal's automated daily run). Refreshing the data just means updating the files in this folder -- roster_refresh
# .csv / seminar_refresh.csv / scholarships_refresh.csv get overwritten in place with the latest pull, and each new
# batch of receipt emails gets saved as its own new thread_<daterange>.csv (never edit or overwrite an existing
# thread_*.csv -- transactions are deduped by Transaction ID when loaded, so old files can just stay put).
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

# Snapshot time this report was built, in the org's local (US Eastern) time -- used to timestamp tabs and to judge
# how recent a registration is, not a live clock. Registration/transaction timestamps in the source sheets are
# Eastern time. Computed live (as of Aug 28, 2026, for the daily automated refresh) rather than hardcoded.
from datetime import datetime
from zoneinfo import ZoneInfo
BUILD_DATE = datetime.now(ZoneInfo("America/New_York")).replace(tzinfo=None)
TODAY_STR = BUILD_DATE.strftime("%b %d, %Y")

wb = openpyxl.Workbook()

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name=FONT, bold=True, size=16, color="1F4E78")
SECTION_FONT = Font(name=FONT, bold=True, size=12, color="1F4E78")
SUB_FONT = Font(name=FONT, size=10, italic=True, color="666666")
BODY_FONT = Font(name=FONT, size=10)
BOLD_FONT = Font(name=FONT, bold=True, size=10)
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
GREEN_FILL = PatternFill("solid", fgColor="D9EAD3")
RED_FILL = PatternFill("solid", fgColor="F4CCCC")
ORANGE_FILL = PatternFill("solid", fgColor="FCE5CD")
GREY_FILL = PatternFill("solid", fgColor="F3F3F3")
BLUE_FILL = PatternFill("solid", fgColor="CFE2F3")  # algorithmically estimated values (vs. yellow = manual/confirmed)
PURPLE_FILL = PatternFill("solid", fgColor="D9D2E9")  # withdrawal-driven Status overrides on the Signups tab (added Oct 6, 2026)
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY = '$#,##0.00'

def style_header(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col+ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

def body_row(ws, row, values, ncols, start_col=1, fills=None):
    for i, v in enumerate(values):
        c = start_col + i
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.alignment = Alignment(vertical="center", wrap_text=False)
        if fills and i in fills:
            cell.fill = fills[i]

def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# =========================================================================
# SHEET 1: Read Me
# =========================================================================
ws = wb.active
ws.title = "Read Me"
ws["B2"] = "Shevet Tapuach -- Payments Tracker"
ws["B2"].font = TITLE_FONT
ws["B3"] = f"Built from Authorize.net receipt emails + the CiviCRM signup roster. Generated Aug 21, 2026; refreshed {TODAY_STR} (this build runs automatically every day)."
ws["B3"].font = SUB_FONT

lines = [
    ("How this sheet is built", "section"),
    ("Every payment your org's payment processor (Authorize.net) emails to tapuach@israelscouts.org / gizbartapuach@israelscouts.org "
     "was pulled from Gmail and logged, one row per transaction, on the Transactions tab. Each one records the amount, date, which "
     "signup it was for (from the receipt's Description field), and the payer's name/email.", "body"),
    ("The Roster tab is a snapshot of 'Yearly Signups for payments' -- a sheet Michal set up with IMPORTRANGE to mirror just the "
     "MembershipReport tab of the main 'Tapuach Yearly Signup 2026-2027' CiviCRM export, pulled at build time. It's used to match a "
     "payment to a specific child by parent email. (Reading the full multi-tab CiviCRM sheet directly proved unreliable -- it has over "
     "a dozen tabs and pulling from it repeatedly dropped or mixed up rows -- so this focused live-linked copy replaced that approach.)", "body"),
    ("The Seminar Roster tab is a snapshot of the similarly-purposed 'Opening Seminar for payments' sheet -- the actual list of kids "
     "registered for the Opening Seminar, with each child's own Participant Fee. Seminar payments are matched against this tab "
     "instead of the main Roster, since it's grade-restricted (10-12 only) and gives an exact per-child fee instead of a guess.", "body"),
    ("The Signups tab groups the raw transactions by (payer + signup) and totals what's been paid so far with a live formula. That's "
     "the view that answers 'who still owes what.'", "body"),
    ("The Scholarships tab is built from Michal's 'Scholarships 2026-2027' sheet (one row per family + event, a % or $ discount off "
     "the normal price). Families are matched by parent name rather than email, and the discounted amount automatically becomes that "
     "family's Expected Total on the Signups tab, so Remaining Balance reflects what they actually owe.", "body"),
    ("The Donations tab lists every Donation Add-on currently logged on the Transactions tab, one row per donation, with the parent's "
     "name and email pulled from the Roster / Seminar Roster (not the raw payer info on the receipt) so it always credits the actual "
     "family on file. Added Sep 4, 2026, per Michal: a storefront order whose item is literally named \"$X Donation\" (a checkout "
     "add-on separate from merchandise) is auto-detected by that item name and logged here too, at whatever amount X is -- it still "
     "shows on Store Purchases as well, since that's genuinely where the payment came through, with a note cross-referencing this tab.",
     "body"),
    ("The Store Purchases tab lists every Tapuach Store (merchandise) charge on its own, since those aren't signup fees and are "
     "otherwise easy to miss mixed in among Annual/Seminar payments on Transactions. Also totaled on the Dashboard. Added Sep 4, "
     "2026, per Michal: it now also includes orders from the new \"Shevet Tapuach\" storefront (a separate platform, Enrollia, "
     "paid via Venmo for now) alongside the older Authorize.net-sourced Tapuach Store charges -- both kinds of charge land in the "
     "same Transactions rows and this same tab, so there's one combined view. The new storefront's order-confirmation emails carry "
     "item-level detail Authorize.net's receipts never did, so this tab gained an Item column (e.g. \"Camp T-Shirt (Size: M * "
     "Color: White)\") and a Participant column (the order's \"For: <name>\" field, when the order was placed on behalf of a "
     "specific child) -- both blank for the older Authorize.net-sourced rows, which never carried that detail. A Participant name "
     "that doesn't match anyone on the Annual or Seminar Roster gets a note flagging it for a spelling check, the same way an "
     "unmatched Venmo signup participant does.", "body"),
    ("The Venmo tab (added Aug 31, 2026, per Michal) lists every Venmo transaction on the org's account since Aug 1, 2026 -- not just "
     "signups, since Venmo is also used day-to-day for unrelated sending/receiving. For each one it shows the other party, the "
     "payer's own comment (or a merchant/event description if they didn't leave one), whether it was incoming or outgoing, and a "
     "signed amount. It's a raw reference ledger and doesn't feed into Signups/Dashboard math -- Venmo signup payments are tracked "
     "separately, see below.", "body"),
    ("Added Aug 31, 2026, per Michal: Venmo signup payments are now tracked on the Transactions tab too. Venmo has no installment "
     "plans (always paid in full) and its automated payment emails carry no participant or event info at all, so instead of trying "
     "to parse that out of Gmail, Michal logs each one herself on a small helper sheet (\"Venmo signups for tracker\": Date, "
     "Participant name, payee name, event, amount). Matched here by the participant's (child's) name against the Roster / Seminar "
     "Roster -- more reliable than trying to match a Venmo display name to a parent -- then merged into Transactions exactly like an "
     "Authorize.net receipt, so it counts toward Signups and Dashboard totals the same way. \"Card / Payment Method\" shows "
     "\"Venmo\" for these rows, and they're exempt from the 2% Authorize.net fee (columns O/P), since Venmo payments never go "
     "through Authorize.net. A participant name that doesn't match anyone on the Roster or Seminar Roster (e.g. a typo) gets "
     "flagged with a note explaining why, rather than silently guessing. One known edge case: a Venmo payment is matched to the "
     "family's Parent 1 email (or Parent 2 if Parent 1 is blank), same as the rest of this sheet's family-level matching -- if that "
     "same family ALSO has an Authorize.net payment for the same signup logged under the other parent's email, the two would show "
     "as two separate Signups rows instead of combining into one family total. Shouldn't come up in the ordinary case (Michal logs "
     "a Venmo entry because that's how the family paid, not as a top-up alongside an existing Authorize.net payment), but worth "
     "knowing about.", "body"),
    ("Added Aug 29, 2026, per Michal: the Reconciliation tab cross-checks Transactions (built from receipt emails) against "
     "Authorize.net's own \"Transaction Detail\" export, which Michal pulls manually from the Authorize.net merchant portal into a "
     "separate Google Sheet (\"Authorize transaction report\") every couple of weeks -- not daily, so it's always somewhat behind. "
     "Any APPROVED transaction in that report (Response Code \"1\") with no matching row on Transactions is a real gap (a receipt "
     "email the Gmail search/parsing missed) and gets listed there for a manual look; a transaction on Transactions with no match "
     "yet in the report is expected (normal lag) and isn't flagged. Revised Oct 5, 2026, per Michal: a declined or errored attempt "
     "(Response Code \"2\"/\"3\") never gets a receipt email at all, since nothing was actually charged -- it still shows up in "
     "Authorize.net's own report, but it correctly has no matching receipt, so it's excluded here rather than flagged as a false "
     "gap. (All 14 \"gaps\" showing before this fix, including the 4 Michal asked about -- Orital Karelic, Gili Sharett, Avishag "
     "Iluz, Amy Nachon -- turned out to be exactly this: a declined attempt, in most cases immediately followed by a successful "
     "retry for the same amount that DOES have a matching receipt.)", "body"),
    ("Added Sep 4, 2026, per Michal: the Payouts tab is a second, different cross-check -- this one from the bank's side rather "
     "than the gateway's. Payments Hub (the org's merchant acquirer/bank processor, separate from Authorize.net the payment "
     "gateway) emails a daily \"Credit/Debit Snapshot\" originally only to tapuach@israelscouts.org; Michal set up a Gmail filter "
     "(Sep 4, 2026) forwarding these to gizbartapuach@israelscouts.org so this pipeline can read them, and separately forwarded "
     "her full backlog back to the start of signups (Aug 10, 2026) so the whole year is covered. Each day's reported Total Batch "
     "Volume is shown next to Authorize.net's own settled-transaction total for that same calendar day (grouped by the Authorize "
     "report's Settlement Date field), but -- after testing this against real overlapping data -- that day-by-day comparison turned "
     "out to be reference only, not a pass/fail: Authorize.net's recorded settlement date routinely lands a day or more off from "
     "the date Payments Hub's batch actually closed once more than one transaction is involved, so a strict same-day match falsely "
     "flagged the large majority of real days as mismatches. What does hold up is a running Cumulative Gap ($ and %) tracked across "
     "the whole overlap window -- it stayed bounded and kept shrinking back down rather than drifting apart (ending around 2.9% of "
     "volume as of the last date both sources cover), which is the actual \"is anything major missing\" signal: a small, "
     "self-correcting gap is normal settlement lag, while a gap that keeps growing and never comes back down would be worth "
     "digging into. Rows beyond the Authorize.net report's own date coverage still show \"Pending\" until Michal refreshes that "
     "sheet. Total Deposit Amount and Month-to-Date Net Deposits (also from the email) are shown for reference only, not forced "
     "into a same-day match either -- real data already shows deposits lag batch volume by an uneven number of days (e.g. Aug 30, "
     "2026 shows $0 deposited that day against $5,420 in batch volume), so a naive same-day \"batch minus fee should equal "
     "deposit\" check would falsely flag every day as a mismatch. One real gap Michal should check for: her forwarded backlog has "
     "no snapshot email for Aug 11 or Aug 12, 2026, even though every other day in the range (including other $0-volume days) has"
     "one -- worth a look in her tapuach inbox (including spam/trash) for those two dates.", "body"),
    ("Added Sep 11, 2026, per Michal: the Payout by Event tab estimates how much of the actual bank-side payout (Payouts tab) is "
     "attributable to each event -- Annual Membership, Opening Seminar, Tapuach Store -- as a cumulative running total by calendar "
     "month starting August 2026 (each month's row includes everything since Aug 1, not just that month alone). It's an estimate: "
     "Payments Hub's own report has no per-event breakdown, so each event's cumulative Authorize.net-settled dollars (classified by "
     "the Invoice Description field, same ANNUAL/SEMINAR/STORE match used on Transactions) is scaled by the ratio of actual "
     "cumulative bank payout $ to cumulative settled $ over the same window -- the same gap the Payouts tab already tracks as "
     "Cumulative Gap. Scope is Authorize.net-settled activity only: Venmo-paid signups/orders never flow through Payments Hub, so "
     "they have no bank payout record to allocate and aren't represented here, even though they show up elsewhere in this workbook. "
     "A second table below the cumulative one (added Sep 11, 2026, per Michal) shows the same breakdown TRANSPOSED -- one column "
     "per month instead of one row -- with each cell showing that month's OWN isolated activity (not cumulative-to-date), so month-"
     "to-month swings are easy to read across a row instead of having to subtract consecutive rows in the cumulative table.",
     "body"),
    ("Added Aug 28, 2026, per Michal: Annual Membership covers every grade, but the Opening Seminar is grades 10-12 only. A "
     "Scholarships row for the Seminar is only shown if the family either has a kid already registered for it, OR has at least one "
     "kid in grades 10-12 on the Annual Roster who could register later. A family matched on the Annual Roster whose kids are all "
     "younger is dropped from the Scholarships tab entirely for that Seminar row, rather than showing a misleading \"not yet "
     "registered\" note -- they're not a candidate yet, not a missing signup. (A family that can't be found on either roster at all "
     "still shows as \"Not yet registered,\" since there's no way to tell whether they're even eligible.)", "body"),
    ("", "body"),
    ("What you need to do", "section"),
    ("Fill in the yellow 'Expected Total' cells on the Signups tab. This sheet can total up what came in, but only your "
     "system (or you) knows the true agreed total for a family that's mid-installment-plan -- so that one number per row is a manual "
     "entry. A handful are already filled in where the amount matched a known price exactly or an even fraction of one (marked with a note).", "body"),
    ("Check any row flagged [!] in the Review Flag column -- it means no matching child could be found on the roster for that payer at "
     "all. (Matching more than one kid isn't flagged, even though \"Matched Child(ren)\" may list several -- Annual and Seminar totals "
     "are always tracked per family, not per kid, so this workbook's math doesn't depend on knowing which specific kid a payment covers.)", "body"),
    ("", "body"),
    ("Keeping it current", "section"),
    ("This is a snapshot, not a live-syncing sheet -- the tools connected in this session can create a new Google Sheet from scratch, "
     "but can't push incremental edits into this file's cells automatically. Two ways to keep it going:", "body"),
    ("  1)  Ask Claude to 'refresh the payments report' periodically (weekly works well) -- it'll pull new receipt emails since "
     "the last build and hand you an updated version.", "body"),
    ("  2)  Add rows yourself: copy a new Authorize.net receipt's Date / Signup / Amount / Payer Name / Payer Email / Card / Invoice / "
     "Transaction ID into a new row on the Transactions tab (below the existing data, within the shaded area). The Signups and Dashboard "
     "tabs' formulas are sized to the actual data every time this sheet is rebuilt (fixed Sep 13, 2026 -- they used to stop at a fixed "
     "row 200, which silently dropped anything added past that row once the season's real transaction count grew past it), so new rows "
     "within the shaded area are picked up automatically on the next refresh -- you'd just need to fill in the Matched Child(ren) and "
     "Review Flag columns by eye, or ask Claude to do a pass over new rows.", "body"),
    ("", "body"),
    ("Known limitations -- please read", "section"),
    ("An earlier version of this sheet had only captured 18 of the ~81 real receipt emails in the account, because "
     "Gmail groups multiple receipts into one email thread and a quick search only previews the ~5 oldest messages per "
     "thread -- the full thread has to be pulled explicitly to see every message. That's what was missing Boaz Gilad's "
     "Aug 16 payments for Ben Gilad (both Annual and Seminar), which is what surfaced this gap. It's now fixed -- all "
     "81 receipts across the 5 threads in the account as of Aug 21 were retrieved and logged below. If a payment still "
     "seems to be missing, it's worth asking Claude to re-check full thread contents rather than assuming a quick "
     "search caught everything.", "body"),
    ("Noa Reikhav's Annual Membership transactions include an apparent duplicate charge on Aug 17 ($207.15 charged "
     "twice within the same second) followed by a $207.15 refund (Credit) on Aug 19 against one of them. All three "
     "rows are logged as-is on Transactions (the refund as a negative amount) and net out correctly to one real "
     "$207.15 payment on the Signups tab -- see the per-row notes on Transactions.", "body"),
    ("Meital (Yitzhak) Sasson's receipts include an Aug 20 \"Tapuach Store\" merchandise charge ($30.00) mixed in "
     "with her Annual and Seminar payments. It's logged on Transactions for a complete record but excluded from the "
     "Signups tab and the Dashboard's by-signup-type totals, since it isn't a signup fee.", "body"),
    ("Authorize.net also sends a separate \"Summary of Automated Recurring Billing\" digest email with two CSV "
     "attachments (Successful.csv / Failed.csv) instead of line-item text -- this session's tools can't open email "
     "attachments, so a recurring-plan charge that only ever showed up in one of those digests (and never generated "
     "its own individual receipt email) wouldn't be captured here. Nothing found while building this version looked "
     "like that case, but it's a blind spot worth knowing about.", "body"),
    ("Authorize.net receipts don't label which installment number a payment is (\"3 of 8\", etc.) -- only the amount and date. "
     "Where a family has made more than one payment for the same signup, this sheet sums them; it doesn't guess how many payments remain "
     "unless the Expected Total is filled in. \"# Payments\" on the Signups tab counts how many have come in so far and updates live as "
     "new rows are added to Transactions; \"Total Installments (if known)\" next to it holds the target count, in two tiers: yellow cells "
     "are confirmed (a family or Michal actually said so -- right now just Brutman/Yonatan, 3 installments, Aug 23, 2026); blue cells are "
     "algorithmically estimated (added Aug 28, 2026, per Michal) by checking whether what a family has paid so far divides evenly into "
     "their full agreed total (net of any scholarship discount, plus any known embedded donation) for some small number of equal "
     "installments -- only kept when the fit is tight, and flagged as an estimate in the Notes column, since it's a pattern match, not "
     "something the family confirmed. Most \"In progress (installment plan)\" rows still end up with neither: they're a single partial "
     "payment with no clean multi-installment pattern to detect, not a family we simply haven't counted. Tell Claude any family's actually "
     "confirmed installment count and it'll get added as a yellow (confirmed) entry on the next refresh.", "body"),
    ("When one parent pays for multiple kids in a single Annual Membership transaction, the receipt doesn't split the amount per "
     "child -- \"Matched Child(ren)\" lists everyone the family's sibling-tier price is based on, not a per-kid breakdown of that "
     "specific charge. That's not flagged as a review item (changed Aug 28, 2026, per Michal) since Annual totals are tracked per "
     "family regardless; if you ever need the actual per-kid split for your own records, ask Claude rather than looking for it here.", "body"),
    ("The Signups tab includes Seminar-registered families with $0 paid so far (Status \"Not yet paid\", shaded red) -- found by "
     "cross-checking every registrant on the Seminar Roster tab against Transactions. The same cross-check hasn't been done for "
     "Annual Membership yet -- the Signups tab there only shows families who've made at least one payment. Ask Claude to run it if "
     "you want the equivalent \"registered but hasn't paid anything\" list for Annual dues.", "body"),
    ("Earlier versions of the Roster tab undercounted signups (72, then 67) because reading the main multi-tab CiviCRM sheet directly "
     "was unreliable -- it silently dropped or duplicated rows depending on which of its many tabs got parsed and how big the pull was. "
     "As of this build, Roster and Seminar Roster are sourced from the two focused '... for payments' sheets instead (158 and 60 rows "
     "respectively as of Aug 28, 2026) and read via a full CSV export rather than a partial text conversion, which resolved it -- including a note that was "
     "previously here claiming Sheves, Rotem's grade disagreed between tabs; both now correctly show 10th Grade, so that was this same "
     "bug, not a real data conflict.", "body"),
    ("Fixed Aug 23, 2026: on the Signups tab, Opening Seminar rows paid by the SECOND parent listed on the Seminar Roster (rather than "
     "the first) were showing a false \"No matching entry on the Seminar Roster tab\" note and, worse, generating a duplicate false "
     "\"Not yet paid\" row for the same family further down -- even though \"Matched Child(ren)\" correctly showed the right kid the whole "
     "time. Root cause: the fee lookup was keyed only by the first parent's email, not both. Affected Polak/Danielle, Gaon/Talia, "
     "Katz/Jonathan, and Wiser/Maya; all four now show correct notes with no duplicate rows.", "body"),
    ("One payer (Yaniv Biran) doesn't have a matching row in the roster -- likely unrelated to the $1.00 test charge below.", "body"),
    ("The $1.00 charge from Yaniv Biran on Aug 10 looks like a card-verification/test charge (common right before a recurring plan "
     "starts), not a real payment -- it's included for completeness but excluded from the Dashboard totals.", "body"),
    ("Both signup forms offer an optional Donation Add-on ($25/$50/$100/$200/$500/$1,000 -- see Price Reference). Receipts don't "
     "separate a donation from the signup fee, so a payment that's more than the expected price may just include one. The "
     "Transactions tab has a \"Donation Add-on\" column (yellow, with a dropdown of the six amounts) -- for anything not covered by the "
     "automatic detection below, log it there yourself, and Total Paid on the Signups tab will automatically exclude it so the balance "
     "still reconciles correctly.", "body"),
    ("Added Aug 28, 2026: Donation Add-ons are now auto-detected and logged for one specific, confirmed pattern -- a family whose "
     "Seminar Roster \"Participant Fee\" isn't a valid standalone price (Price Reference only lists $345 or $305 per participant) "
     "because the registration form folded a voluntary donation straight into that fee instead of recording it separately. When the "
     "fee minus one of the six donation tiers lands back on a valid price, the difference is treated as a donation: Expected Total on "
     "Signups uses the real base fee, and the donation itself is logged on the family's most recent Transactions row for that signup as "
     "soon as there's any payment -- regardless of whether more installments are still due (with a note explaining why). This is what "
     "fixed Packin/Hadas (roster shows $330 -- really $305 + a $25 donation, now correctly shown paid in full with a $25 donation logged) "
     "and Brutman/Yonatan (roster shows $370 -- really $345 + a $25 donation via a 3-installment plan Michal confirmed on Aug 23, 2026; "
     "the $25 is logged against installment 1 even though only $123.34 of the $345 base has come in so far -- per Michal, Aug 28, 2026, "
     "the donation shouldn't wait on the remaining installments). Other families with the same roster-fee pattern as "
     "of this build: Torjman Thomas/Josef and Preis/Noga ($330 = $305 + $25, both already paid in full and logged), and Mosca/Hailey "
     "($405 = $305 + $100, already paid in full and logged). This detection is scoped to the Seminar Roster's fee field specifically -- a "
     "generic \"any Annual overpayment might be a donation\" check was NOT built in, since a scan of current Annual payments found no "
     "case where a family's total payments overshoot the sibling-tier price by a donation-tier amount (every Annual family currently "
     "shown as short is genuinely still mid-installment-plan, not carrying a hidden donation).", "body"),
    ("Added Oct 3, 2026 (automated daily refresh): the storefront (info@enrollia.graysbow.com) sent a THIRD email template not "
     "previously seen -- \"Order #N Status Update - Shevet Tapuach\", a post-purchase fulfillment/cancellation notice (distinct from "
     "the \"Order Confirmation #N\"/\"Order Confirmed\" payment-confirmation templates STEP 4c parses). Seven arrived on Oct 2, 2026 "
     "(orders #2, #6, #7, #8, #9, #10, #11). None were parsed into store_orders_*.csv or this workbook -- they carry a fulfillment/"
     "cancellation status, not a \"Payment Status: Paid\" confirmation, so they don't fit the existing schema and the default choice"
     "was not to guess. Orders #9, #10, and #11 (marked \"fulfilled\") already match orders captured as paid back on Sep 2, 2026 -- no "
     "action needed there. Orders #2, #6, and #7 (also \"fulfilled\") reference Tal/Omer/Gil Sheves purchases that do NOT appear "
     "anywhere in this workbook's captured data, and order #8 (marked \"cancelled\" -- Camp T-Shirt, For: Omer Sheves, $1.00) also has "
     "no matching captured paiment. Worth Michal's review in the storefront's own admin panel: were #2/6/7/8 ever actually paid (and"
     "if so, their original confirmation emails were missed by every prior run too), or were they test/void orders that never took a "
     "real payment? If Michal confirms a pattern for how to handle this status-update template going forward, it can be added here.",
     "body"),
    ("Added Oct 3, 2026 (automated daily refresh): confirmed a real, concerning limitation in the build.py Drive-backup mechanism "
     "(STEP 8/0) -- the Google Drive tool available to this pipeline has no \"upload from local file path\" option, so every one of "
     "the 31 build_py.plain.part*of31 chunks has to be retyped, character-by-character, as a literal tool-call parameter. This run "
     "found build_py.plain.part31of31 genuinely corrupted on Drive (one word altered to another same-length word), which a byte-size "
     "check alone did NOT catch -- only a full byte-for-byte diff against the locally-reconstructed canonical chunk did. Re-uploading "
     "it correctly took three attempts, because the SAME retyping step used to fix it, and separately to verify it by re-downloading, "
     "kept introducing its own fresh one-off errors (a dropped character, a stray byte, a wrong cell-reference column letter). After the "
     "fix, parts 16, 20, 21, and 29 (the ones a prior session's subagent had flagged as needing multiple self-correction rounds) were "
     "each independently re-downloaded and diffed byte-for-byte against the canonical local chunks and confirmed correct; all 31 parts "
     "match their canonical chunks' byte sizes exactly. Given how consistently this exact retyping step (not the underlying logic) has "
     "been the source of error on both the upload and the verification side, the remaining 26 parts were NOT individually re-verified "
     "byte-for-byte this run -- doing so with the same error-prone method would add more risk than it removes. This is a structural gap "
     "worth raising with Michal directly: a backup mechanism that doesn't require an LLM to retype long opaque text (e.g. a tool that "
     "accepts a local file path, or storing build.py in a form Drive can diff/checksum directly) would close it for good.", "body"),
    ("Added Oct 4, 2026 (automated daily refresh) -- defensible-default fix for a newly-discovered data anomaly, not a silent one: "
     "the live \"Scholarships 2026-2027\" sheet's Type column (Percentage vs. flat $) now reads \"Perentage\" (a typo) for every row, "
     "which would have made this build crash outright (it no longer matched the code's percentage check, so every non-blank discount "
     "fell into the flat-dollar-amount branch and tried to parse a value like \"100.00%\" as a plain number). Rather than guess at or "
     "hand-correct the sheet's typo, the code now also treats a Discount value containing an actual \"%\" character as a percentage "
     "regardless of what the Type column says -- the same way a person skimming the column would read it. The original Type-text check "
     "is kept alongside it, so nothing about correctly-spelled rows changed. Worth a glance at the Scholarships 2026-2027 sheet's Type "
     "dropdown when convenient, but no action is required -- this build reads it correctly either way now.", "body"),
    ("Added Oct 4, 2026 (automated daily refresh) -- this build would not have run at all without a fix: the Payouts tab's "
     "payout_rows_raw loading loop referenced a variable name (`_po_rows`) that was never assigned anywhere in the file -- only "
     "`_por_rows` (one line above it) was. That's a one-character drop, and since at least one payouts_*.csv file has existed "
     "since August, this exact line would raise a NameError on every run, every day -- so either this session's mechanical "
     "build.py reconstruction from the Drive-stored text parts introduced the drop, or it was already sitting in the Drive "
     "backup (possibly from an earlier day's reconstruction/upload). Which of the two isn't something this run can tell for "
     "certain. Either way: fixed by using the only variable that actually exists (`_por_rows`), not a guess at new logic, and the "
     "corrected build.py has been re-pushed to the Drive backup folder so tomorrow's run doesn't inherit the same break.", "body"),
    ("Added Oct 4, 2026 (automated daily refresh): a new Authorize.net receipt came in with Invoice Description "
     "\"Shevet Tapuach - Yom Tzofe 2026\" (Keshet Givol, $181.25, and a second one, $190.00 from Michal Sheves) -- a payment page "
     "name never seen before, matching neither the ANNUAL nor SEMINAR description strings. The existing fallback (any Invoice "
     "Description that doesn't match Annual Membership or Opening Seminar is bucketed as Tapuach Store merchandise) applied "
     "automatically, so this transaction was counted as a Store Purchase rather than left out -- the defensible default already "
     "built into this pipeline, not a new guess. RESOLVED Oct 5, 2026 (see next note): Yom Tzofe now has its own category.", "body"),
    ("Added Oct 5, 2026, per Michal: Yom Tzofe 2026 is now tracked as its own event, the same way Opening Seminar is -- its own "
     "Google Sheet (\"Yom Tzofe for payments\"), its own \"Yom Tzofe Roster\" tab here, its own line on Price Reference, and its "
     "own classification on Transactions/Signups/Dashboard/Payout by Event. Simpler than Seminar in three ways Michal confirmed: "
     "a flat $95/participant fee (no sibling discount), no grade restriction (everyone can sign up), and no scholarships. Note: "
     "re-checking the receipt flagged above once the underlying thread_*.csv data was back in sync (Oct 5), only the Michal Sheves "
     "$190.00 charge actually carries the Yom Tzofe description -- the Keshet Givol $181.25 charge reads \"Annual Membership\" both "
     "times it appears (Sep 5 and Oct 4), so the prior note's claim that both were Yom Tzofe looks mistaken. Michal Sheves' $190.00 "
     "correctly reclassifies from Tapuach Store to Yom Tzofe on this and future runs -- that amount doesn't divide evenly by $95 "
     "(likely 2 participants plus a partial amount, or a donation folded in), so it's worth a glance on the Signups tab rather than "
     "assumed correct.", "body"),
    ("Added later on Oct 5, 2026, per Michal: Yom Tzofe scholarships are now supported too -- she added a third (\"Yom Tzofe\") "
     "section to the \"Scholarships 2026-2027\" sheet, matched and discounted the same way Annual/Seminar scholarships already "
     "are (by parent name against the Yom Tzofe Roster, reducing the flat $95/participant total by the entered % or $ amount), "
     "and shown on the Scholarships tab alongside the other two events. The earlier \"no scholarships for this event\" notes above "
     "and on Price Reference/Yom Tzofe Roster describe how the event launched that morning, not how it works now.", "body"),
    ("Added Oct 6, 2026, per Michal: a new Withdrawals tab tracks families who withdraw a participant mid-season, sourced from "
     "a small helper sheet she maintains, \"Withdrawals for Payments tracker\" (Date, parent name, participant name, event, "
     "refund type [Refund/Adjust/Both], updated installment amount, refund amount). A withdrawal can mean a (usually partial) "
     "REFUND through Authorize.net and/or an ADJUST to the remaining installment plan when a sibling who isn't withdrawing still "
     "owes their own fee. Per Michal: don't assume a flat \"$30 fee\" off the refund -- the Withdrawals tab instead cross-checks "
     "the logged refund amount against the ACTUAL Authorize.net credit receipt(s) on the Transactions tab (a refund can be split "
     "across more than one credit, since a single credit can't exceed the original charge) and goes by that real total. "
     "\"Updated installment amount\" is informational only -- the new per-installment dollar figure Michal sets directly in "
     "Authorize.net for the remaining sibling's plan (the installment count doesn't change); it doesn't override Expected Total "
     "here. The matching family+event row on the Signups tab gets its Status flipped to \"Refund (withdrawal)\", \"Adjusted "
     "(withdrawal)\", or \"Refund + Adjusted (withdrawal)\" (shaded purple) with the same details folded into its Notes -- this "
     "replaces the usual live paid-in-full/in-progress formula for that row going forward, by design (a withdrawal is treated as "
     "a one-time flag, not an ongoing payment state). Removing a withdrawn participant from the live Roster/Seminar Roster/Yom "
     "Tzofe Roster sheet is Michal's own step, not something this build does -- the existing sibling-tier pricing already "
     "recomputes automatically off however many kids are left there. One consequence worth knowing: because the roster-match "
     "check on Transactions is recomputed live against the CURRENT roster every run, a family that withdraws ENTIRELY (removed "
     "from the roster completely) would otherwise make their own past, already-fine transactions flip to a false \"No roster "
     "match found\" -- the Transactions build now checks the Withdrawals helper sheet for exactly this case and substitutes a "
     "withdrawal note instead of a false flag, matched by payer name (not email, since the family's email may no longer be on "
     "the roster at all to resolve).", "body"),
]
r = 5
for text, kind in lines:
    cell = ws.cell(row=r, column=2, value=text)
    if kind == "section":
        cell.font = SECTION_FONT
        r += 1
    else:
        cell.font = BODY_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28 if text else 8
        r += 1
autofit(ws, [3, 118])
ws.sheet_view.showGridLines = False

# =========================================================================
# SHEET 2: Price Reference
# =========================================================================
ws = wb.create_sheet("Price Reference")
ws["B2"] = "Price Reference"
ws["B2"].font = TITLE_FONT
ws["B3"] = "Pulled from the two registration-setup forms submitted this year. Used as a lookup aid on the Signups tab."
ws["B3"].font = SUB_FONT

ws["B5"] = "Annual Membership 2026-27 -- sibling-discount pricing (per Michal, Aug 21, 2026)"
ws["B5"].font = SECTION_FONT
headers = ["Tier", "Applies", "1st Kid", "2nd Kid", "3rd Kid", "4th Kid"]
for i, h in enumerate(headers):
    ws.cell(row=6, column=2+i, value=h)
style_header(ws, 6, len(headers), start_col=2)
rows = [
    ["Early Bird", "Through Sep 24, 2026", 740, 710, 680, 650],
    ["Regular", "Sep 25, 2026 onward", 790, 760, 730, None],
]
for i, row in enumerate(rows):
    body_row(ws, 7+i, row, len(row), start_col=2)
    for c in range(4, 8):
        ws.cell(row=7+i, column=c).number_format = CURRENCY
ws["B10"] = ("Price is per family based on how many kids are on the Annual Membership, not per kid individually -- e.g. a family "
             "with 2 kids registering Early Bird owes $740 + $710 = $1,450 total, not $740 x 2. All transactions logged so far "
             "fall under Early Bird pricing (through Sep 24, 2026).")
ws["B10"].font = SUB_FONT
ws["B10"].alignment = Alignment(wrap_text=True)
ws.merge_cells("B10:G11")
ws.row_dimensions[10].height = 40
ws["B12"] = ("Note: the Regular-tier 4th-kid price wasn't given -- if the same +$50-per-kid pattern holds it would be $700, "
             "but please confirm before relying on it. Also unconfirmed: pricing for 5+ kids in one family.")
ws["B12"].font = SUB_FONT
ws["B12"].alignment = Alignment(wrap_text=True)
ws.merge_cells("B12:G13")
ws.row_dimensions[12].height = 40

ws["B16"] = "Opening Seminar (Sep 4-6, 2026, Camp Kennybrook)"
ws["B16"].font = SECTION_FONT
headers2 = ["Category", "Price", "", "", "", ""]
for i, h in enumerate(headers2[:2]):
    ws.cell(row=17, column=2+i, value=h)
style_header(ws, 17, 2, start_col=2)
seminar_rows = [
    ["Shachbag Regular (grades 10-12)", 305],
    ["Tzevet Movil (Leading Team)", 345],
]
for i, row in enumerate(seminar_rows):
    body_row(ws, 18+i, row, len(row), start_col=2)
    ws.cell(row=18+i, column=3).number_format = CURRENCY
ws["B21"] = "Up to 3 installments allowed. Only registered annual members in grades 10-12 attend."
ws["B21"].font = SUB_FONT

ws["B24"] = "Donation Add-on (optional, on top of the signup price)"
ws["B24"].font = SECTION_FONT
ws["B25"] = ("Both the Annual and Seminar signup forms offer an optional donation checkbox at fixed amounts. When checked, it's "
             "added straight into the total charged -- which is why some payments (or installments) don't divide evenly into a "
             "clean multiple of the signup price. If a payment looks off by one of these amounts, check the Transactions Donation "
             "Add-on column.")
ws["B25"].font = SUB_FONT
ws["B25"].alignment = Alignment(wrap_text=True)
ws.merge_cells("B25:G26")
ws.row_dimensions[25].height = 40
donation_tiers = [25, 50, 100, 200, 500, 1000]
ws.cell(row=28, column=2, value="Donation #").font = HEADER_FONT
ws.cell(row=28, column=2).fill = HEADER_FILL
ws.cell(row=28, column=3, value="Amount").font = HEADER_FONT
ws.cell(row=28, column=3).fill = HEADER_FILL
style_header(ws, 28, 2, start_col=2)
for i, amt in enumerate(donation_tiers):
    body_row(ws, 29+i, [f"Donation #{i+1}", amt], 2, start_col=2)
    ws.cell(row=29+i, column=3).number_format = CURRENCY

ws["B37"] = "Yom Tzofe 2026 (added Oct 5, 2026, per Michal)"
ws["B37"].font = SECTION_FONT
ws.cell(row=38, column=2, value="Category").font = HEADER_FONT
ws.cell(row=38, column=2).fill = HEADER_FILL
ws.cell(row=38, column=3, value="Price").font = HEADER_FONT
ws.cell(row=38, column=3).fill = HEADER_FILL
style_header(ws, 38, 2, start_col=2)
body_row(ws, 39, ["Everyone (flat fee, no sibling discount)", 95], 2, start_col=2)
ws.cell(row=39, column=3).number_format = CURRENCY
ws["B40"] = "No grade restriction (everyone can sign up). Scholarships ARE supported (added later on Oct 5, 2026, per Michal) -- see the Scholarships tab."
ws["B40"].font = SUB_FONT

ws["B43"] = "Sources"
ws["B43"].font = SECTION_FONT
ws["B44"] = "Annual pricing (sibling-discount tiers above): provided directly by Michal, Aug 21, 2026 -- supersedes the flat per-kid tiers in the original registration form, which didn't match observed payments."
ws["B44"].font = BODY_FONT
ws["B45"] = "Seminar pricing: Gmail -- \"Thanks for filling out this form: Camp/Event/Activity Registration Form\", submitted Aug 10, 2026."
ws["B45"].font = BODY_FONT
ws["B46"] = "Donation Add-on tiers: provided directly by Michal, Aug 21, 2026 (screenshot of the registration form's donation checkboxes)."
ws["B46"].font = BODY_FONT
ws["B47"] = "Yom Tzofe pricing ($95 flat fee, everyone, no grade restriction): provided directly by Michal, Oct 5, 2026. Scholarships for this event added later the same day."
ws["B47"].font = BODY_FONT
autofit(ws, [3, 34, 14, 12, 12, 12])
ws.sheet_view.showGridLines = False

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("sheets 1-2 done")

# =========================================================================
# SHEET 3: Roster (snapshot)
# =========================================================================
ws = wb.create_sheet("Roster")
ws["A1"] = (f"Roster snapshot -- from the \"Yearly Signups for payments\" sheet, a live IMPORTRANGE feed of the MembershipReport tab on "
            f"\"Tapuach Yearly Signup 2026-2027\", pulled {TODAY_STR} (this report refreshes automatically every day).")
ws["A1"].font = SUB_FONT
ws.merge_cells("A1:I1")
ws.row_dimensions[1].height = 18

headers = ["Child Name", "Grade", "Membership Type", "Status", "Child Email",
           "Parent 1 Name", "Parent 1 Email", "Parent 2 Name", "Parent 2 Email"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

with open(os.path.join(DATA_DIR, "roster_refresh.csv")) as f:
    reader = csv.DictReader(f)
    roster_rows = list(reader)

for i, rrow in enumerate(roster_rows):
    vals = [rrow["Contact Name"], rrow["Grade This Year"], rrow["Membership Type"], rrow["Status"],
            rrow["Email"], rrow["Parent 1 Contact Name"], rrow["Parent 1 Email"],
            rrow["Parent 2 Contact Name"], rrow["Parent 2 Email"]]
    body_row(ws, 3+i, vals, len(vals), start_col=1)

ws.freeze_panes = "A3"
autofit(ws, [26, 11, 30, 10, 26, 20, 26, 20, 26])
ws.auto_filter.ref = f"A2:I{2+len(roster_rows)}"
print("roster rows written:", len(roster_rows))

wb.save("/home/claude/tapuach_report/scaffold.xlsx")

# =========================================================================
# SHEET 3b: Seminar Roster (snapshot) -- who's actually registered, and their exact fee
# =========================================================================
ws = wb.create_sheet("Seminar Roster")
ws["A1"] = (f"Registrant snapshot -- from the \"Opening Seminar for payments\" sheet, a live IMPORTRANGE feed of the Opening Seminar "
            f"registration tab, pulled {TODAY_STR} (this report refreshes automatically every day). Participant Fee is the system's own "
            f"per-child charge -- use it as the source of truth for what's owed, rather than reverse-engineering it from payment amounts.")
ws["A1"].font = SUB_FONT
ws.merge_cells("A1:G1")
ws.row_dimensions[1].height = 28

headers = ["Participant Name", "Grade", "Participant Fee", "Registration Date", "Parent 1 Email", "Parent 2 Email"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

with open(os.path.join(DATA_DIR, "seminar_refresh.csv")) as f:
    reader = csv.DictReader(f)
    seminar_roster_rows = list(reader)

for i, srow in enumerate(seminar_roster_rows):
    fee = float(srow["Participant Fee"].replace("$", "").replace(",", ""))
    vals = [srow["Participant Name"], srow["Grade This Year"], fee, srow["Registration Date"],
            srow["Parent 1 Email"], srow["Parent 2 Email"]]
    body_row(ws, 3+i, vals, len(vals), start_col=1)
    ws.cell(row=3+i, column=3).number_format = CURRENCY

ws.freeze_panes = "A3"
autofit(ws, [30, 11, 13, 18, 28, 28])
ws.auto_filter.ref = f"A2:F{4+len(seminar_roster_rows)}"
print("seminar roster rows written:", len(seminar_roster_rows))

wb.save("/home/claude/tapuach_report/scaffold.xlsx")

# Email -> matched Seminar registrants (name + grade), and per-family Seminar fee total.
email_to_seminar_kids = defaultdict(list)
_seminar_family_totals = defaultdict(float)  # summed once per family, keyed by a canonical family key
_seminar_family_kids_count = defaultdict(int)
for srow in seminar_roster_rows:
    fee = float(srow["Participant Fee"].replace("$", "").replace(",", ""))
    for email_col in ("Parent 1 Email", "Parent 2 Email"):
        e = srow[email_col].strip().lower().replace("\\_", "_")
        if e:
            email_to_seminar_kids[e].append(srow["Participant Name"] + " (" + srow["Grade This Year"] + ")")
    # attribute fee once per family, keyed by Parent 1 Email (or Parent 2 if P1 blank)
    key = (srow["Parent 1 Email"].strip().lower().replace("\\_", "_")
           or srow["Parent 2 Email"].strip().lower().replace("\\_", "_"))
    _seminar_family_totals[key] += fee
    _seminar_family_kids_count[key] += 1

# Some families' Seminar Roster "Participant Fee" bakes in a voluntary Donation Add-on instead of keeping it
# separate -- the registration form just charges (and records) base fee + donation as one number. Detected
# automatically (Aug 28, 2026, after Michal flagged Packin/Hadas and re-flagged Brutman/Yonatan) by checking
# whether the family's roster total fails to match any valid combination of the two real per-participant prices
# ($305/$345 -- see Price Reference) but DOES match one once a single Donation Add-on tier ($25/$50/$100/$200/
# $500/$1,000) is subtracted back out. Where that's the case, seminar_family_fee below is set to the real base
# fee (excluding the donation) and the donation itself is auto-logged on the Transactions tab once the family's
# payments reach that base fee -- see the donation auto-detection block after the Transactions loop.
DONATION_TIERS = [25, 50, 100, 200, 500, 1000]
SEMINAR_BASE_PRICES = [305, 345]

def _decompose_seminar_fee(total, n_kids):
    """(base_total, donation_amt) -- donation_amt is 0 if `total` already matches n_kids valid Seminar prices."""
    from itertools import product
    valid_combos = set(sum(c) for c in product(SEMINAR_BASE_PRICES, repeat=n_kids)) if n_kids else {0}
    if round(total, 2) in valid_combos:
        return total, 0.0
    for tier in DONATION_TIERS:
        if round(total - tier, 2) in valid_combos:
            return total - tier, float(tier)
    return total, 0.0  # doesn't decompose cleanly -- leave the roster figure as-is, no donation assumed

_seminar_family_decomp = {key: _decompose_seminar_fee(total, _seminar_family_kids_count[key])
                           for key, total in _seminar_family_totals.items()}

# seminar_family_fee: looked up by WHICHEVER parent email actually shows up as a payer on a transaction (either
# parent may pay), so it needs to resolve to the same family total regardless of which parent's email is used --
# not just the canonical Parent-1-first key above. (Bug found Aug 23, 2026: Danielle Polak's Seminar payment came
# from Parent 2's email, liron_polak@yahoo.com, but the fee total was only keyed under Parent 1's orit_polak@yahoo.com,
# so the payment showed a false "No matching entry on the Seminar Roster tab" note even though "Matched Child(ren)"
# correctly showed her -- and a duplicate false "not yet paid" row was added below for the same family. Also affected
# Gaon/Talia, Katz/Jonathan, and Wiser/Maya, all paid by Parent 2.)
seminar_family_fee = {}
seminar_family_donation = {}  # email -> (donation_amt, base_total, raw_roster_total) for embedded-donation families
for srow in seminar_roster_rows:
    key = (srow["Parent 1 Email"].strip().lower().replace("\\_", "_")
           or srow["Parent 2 Email"].strip().lower().replace("\\_", "_"))
    raw_total = _seminar_family_totals.get(key, 0.0)
    base_total, donation_amt = _seminar_family_decomp.get(key, (raw_total, 0.0))
    for email_col in ("Parent 1 Email", "Parent 2 Email"):
        e = srow[email_col].strip().lower().replace("\\_", "_")
        if e:
            seminar_family_fee[e] = base_total
            if donation_amt:
                seminar_family_donation[e] = (donation_amt, base_total, raw_total)

# =========================================================================
# SHEET 3c: Yom Tzofe Roster (snapshot) -- added Oct 5, 2026, per Michal: a new one-day event with its own
# registration sheet ("Yom Tzofe for payments"), parallel to how Seminar has its own roster. Unlike Seminar,
# there's no sibling discount, no grade restriction (everyone can sign up), and no folded-in donation to
# decompose -- it's a flat $95 per participant (confirmed by Michal, Oct 5, 2026), so this section is simpler
# than the Seminar Roster block above. The sheet also has only ONE parent-email column ("Parent 1 Email"), not
# two, with "Email" (the registrant's own address, which may be the kid's own email) as the fallback family key
# when Parent 1 Email is blank. Scholarships ARE supported here (added later the same day, per Michal) --
# applied the same way as Annual/Seminar, see the Scholarships tab/scholarship_expected below.
# =========================================================================
ws = wb.create_sheet("Yom Tzofe Roster")
ws["A1"] = (f"Registrant snapshot -- from the \"Yom Tzofe for payments\" sheet, a live IMPORTRANGE feed of the Yom Tzofe "
            f"registration tab, pulled {TODAY_STR} (this report refreshes automatically every day). Flat fee is $95/participant "
            f"(no sibling discount, no grade restriction, no scholarships) -- see Price Reference.")
ws["A1"].font = SUB_FONT
ws.merge_cells("A1:F1")
ws.row_dimensions[1].height = 28

headers = ["Participant Name", "Grade", "Fee", "Registration Date", "Email", "Parent 1 Email"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

with open(os.path.join(DATA_DIR, "yomtzofe_refresh.csv")) as f:
    reader = csv.DictReader(f)
    yomtzofe_roster_rows = list(reader)

YOM_TZOFE_FEE = 95.0  # flat per-participant fee, everyone eligible -- confirmed by Michal, Oct 5, 2026

for i, yrow in enumerate(yomtzofe_roster_rows):
    vals = [yrow["Participant Name"], yrow["Grade This Year"], YOM_TZOFE_FEE, yrow["Registration Date"],
            yrow["Email"], yrow["Parent 1 Email"]]
    body_row(ws, 3+i, vals, len(vals), start_col=1)
    ws.cell(row=3+i, column=3).number_format = CURRENCY

ws.freeze_panes = "A3"
autofit(ws, [30, 11, 10, 18, 28, 28])
ws.auto_filter.ref = f"A2:F{4+len(yomtzofe_roster_rows)}"
print("yom tzofe roster rows written:", len(yomtzofe_roster_rows))

wb.save("/home/claude/tapuach_report/scaffold.xlsx")

# Email -> matched Yom Tzofe registrants (name + grade), and per-family flat-fee total ($95 x however many
# participants share that family's canonical key). Family key is Parent 1 Email, falling back to the
# registrant's own Email column when Parent 1 Email is blank (this sheet has no Parent 2 Email column at all).
email_to_yomtzofe_kids = defaultdict(list)
_yomtzofe_family_kids_count = defaultdict(int)
for yrow in yomtzofe_roster_rows:
    _key = (yrow["Parent 1 Email"].strip().lower().replace("\\_", "_")
            or yrow["Email"].strip().lower().replace("\\_", "_"))
    for _e_raw in (yrow["Parent 1 Email"], yrow["Email"]):
        _e = _e_raw.strip().lower().replace("\\_", "_")
        if _e:
            email_to_yomtzofe_kids[_e].append(yrow["Participant Name"] + " (" + yrow["Grade This Year"] + ")")
    _yomtzofe_family_kids_count[_key] += 1

yomtzofe_family_fee = {}
for yrow in yomtzofe_roster_rows:
    _key = (yrow["Parent 1 Email"].strip().lower().replace("\\_", "_")
            or yrow["Email"].strip().lower().replace("\\_", "_"))
    _total = YOM_TZOFE_FEE * _yomtzofe_family_kids_count.get(_key, 0)
    for _e_raw in (yrow["Parent 1 Email"], yrow["Email"]):
        _e = _e_raw.strip().lower().replace("\\_", "_")
        if _e:
            yomtzofe_family_fee[_e] = _total

# =========================================================================
# SHEET 4: Transactions
# =========================================================================
ANNUAL = "Shevet Tapuach - Annual Membership Payment Page 2026-27"
SEMINAR = "Shevet Tapuach - Opening Seminar 2026-27"
YOM_TZOFE = "Shevet Tapuach - Yom Tzofe 2026"
STORE = "Tapuach Store (merchandise -- not a signup fee)"

# Participant (child) name -> family email. Used ONLY to match Venmo-paid signups (see the Venmo block below) --
# Michal logs the participant's actual name on her "Venmo signups for tracker" helper sheet, which is far more
# reliable than trying to fuzzy-match a Venmo display name to a parent's roster name (the approach used for
# Authorize.net's name-fallback matching above/below).
child_name_to_family_email = {}
for rr in roster_rows:
    _child_key = rr["Contact Name"].strip().lower()
    _fam_email = rr["Parent 1 Email"].strip().lower() or rr["Parent 2 Email"].strip().lower() or rr["Email"].strip().lower()
    if _child_key and _fam_email:
        child_name_to_family_email.setdefault(_child_key, _fam_email)

child_name_to_seminar_family_email = {}
for srow in seminar_roster_rows:
    _child_key = srow["Participant Name"].strip().lower()
    _fam_email = (srow["Parent 1 Email"].strip().lower().replace("\\_", "_")
                  or srow["Parent 2 Email"].strip().lower().replace("\\_", "_"))
    if _child_key and _fam_email:
        child_name_to_seminar_family_email.setdefault(_child_key, _fam_email)

child_name_to_yomtzofe_family_email = {}
for yrow in yomtzofe_roster_rows:
    _child_key = yrow["Participant Name"].strip().lower()
    _fam_email = (yrow["Parent 1 Email"].strip().lower().replace("\\_", "_")
                  or yrow["Email"].strip().lower().replace("\\_", "_"))
    if _child_key and _fam_email:
        child_name_to_yomtzofe_family_email.setdefault(_child_key, _fam_email)

# Full transaction set. Originally rebuilt Aug 21, 2026 after discovering that an earlier pass had only captured
# 18 of the ~81 real receipt emails (Gmail groups receipts into threads and a quick search only previews the ~5
# oldest messages per thread -- the full thread had to be pulled explicitly to see everything). Staged as CSVs from
# the "Merchant Email Receipt" Gmail threads, then merged and deduped by Transaction ID. Refreshed Aug 23, 2026
# with a new thread of receipts from Aug 22-23, and again Aug 28, 2026 with a thread covering Aug 24-28 (58 more
# transactions, pulled via full get_thread rather than the truncated search preview).
#
# Loads EVERY data/thread_*.csv file rather than a hardcoded list (changed Aug 28, 2026 when daily automated
# refreshes were set up) -- each new batch of receipts just needs to be saved as its own new thread_<daterange>.csv
# in that folder; dedup by Transaction ID below means old files never need to be touched or merged by hand.
import re
from datetime import datetime as _dt

def _load_tx_csv(path):
    with open(path) as f:
        return list(csv.DictReader(f))

_raw_tx = []
for _path in sorted(glob.glob(os.path.join(DATA_DIR, "thread_*.csv"))):
    _raw_tx += _load_tx_csv(_path)

_seen_txnid = set()
_dedup_tx = []
for _r in _raw_tx:
    if _r["txnid"] in _seen_txnid:
        continue
    _seen_txnid.add(_r["txnid"])
    _dedup_tx.append(_r)

def _parse_dt(s):
    s = re.sub(r"\s+[A-Z]{2,4}$", "", s.strip())  # strip trailing timezone abbreviation (EDT)
    return _dt.strptime(s, "%d-%b-%Y %H:%M:%S")

def _parse_card(method):
    return method.strip().replace("American Express", "Amex")

_parsed_tx = []
for _r in _dedup_tx:
    _d = _parse_dt(_r["datetime"])
    _desc = _r["desc"].strip()
    # Substring match (not exact-equality) because Authorize.net sometimes wraps the normal description in extra
    # framing text -- e.g. "Offline Registration for Event: Shevet Tapuach - Opening Seminar 2026-27 by: Michal
    # Sheves" for a manually-keyed offline payment. That still contains the SEMINAR string verbatim, so it's
    # correctly bucketed as a Seminar signup instead of falling through to STORE. YOM_TZOFE added Oct 5, 2026 --
    # until now, any "Shevet Tapuach - Yom Tzofe 2026" receipt fell through to STORE (see Read Me).
    _signup = (ANNUAL if ANNUAL in _desc else
               (SEMINAR if SEMINAR in _desc else
                (YOM_TZOFE if YOM_TZOFE in _desc else STORE)))
    _amount = float(_r["amount"])
    if _r["txtype"].strip() == "Credit":
        _amount = -_amount
    _row = (_d.strftime("%Y-%m-%d"), _d.strftime("%I:%M %p"), _signup, _amount,
            _r["name"].strip(), _r["email"].strip().lower(), _parse_card(_r["method"]),
            _r["invoice"].strip(), _r["txnid"].strip())
    _parsed_tx.append((_d, _row))
_parsed_tx.sort(key=lambda x: x[0])

# =========================================================================
# Venmo-paid signups -- added Aug 31, 2026, per Michal. Venmo has no installment plans (always pay in full) and
# its automated payment emails carry no participant/event info at all, so rather than parsing that out of Gmail,
# Michal logs each Venmo signup payment herself on a small helper sheet ("Venmo signups for tracker": Date,
# Participant name, payee name [the payer's Venmo display name], event, amount) -- pulled into
# data/venmo_signups.csv each refresh, same optional-file treatment as authorize_report.csv. Matched below by
# PARTICIPANT name (the child_name_to_*_family_email maps above) rather than payer email/name, then merged
# straight into the same `transactions` list Authorize.net receipts use, so every downstream step (roster-match
# flag, Signups rollup, Dashboard totals) treats a Venmo payment exactly like an Authorize.net one -- except the
# 2% processing fee, which only applies to Authorize.net charges (see the Transactions sheet build below).
_VENMO_SIGNUPS_PATH = os.path.join(DATA_DIR, "venmo_signups.csv")
venmo_notes = {}  # synthetic Venmo txn_id -> note, merged into txn_notes once that dict exists below
_venmo_parsed = []
if os.path.exists(_VENMO_SIGNUPS_PATH):
    with open(_VENMO_SIGNUPS_PATH, newline="") as f:
        _venmo_rows = list(csv.DictReader(f))
    for _vr in _venmo_rows:
        _participant = (_vr.get("Participant name") or "").strip()
        _payee = (_vr.get("payee name") or "").strip()
        _event_raw = (_vr.get("event") or "").strip()
        _amt_raw = (_vr.get("amount") or "").strip()
        _date_raw = (_vr.get("Date") or "").strip()
        if not _participant or not _amt_raw:
            continue  # blank/incomplete row -- skip rather than guess
        try:
            _v_amount = float(_amt_raw.replace("$", "").replace(",", ""))
        except ValueError:
            continue
        try:
            _v_dt = _dt.strptime(_date_raw, "%Y-%m-%d")
        except ValueError:
            try:
                _v_dt = _dt.strptime(_date_raw, "%m/%d/%Y")
            except ValueError:
                _v_dt = BUILD_DATE  # unparseable/blank date -- fall back to today rather than dropping the row
        _event_key = _event_raw.lower()
        if "annual" in _event_key:
            _v_signup = ANNUAL
        elif "seminar" in _event_key:
            _v_signup = SEMINAR
        elif "yom tzofe" in _event_key or "yomtzofe" in _event_key:
            _v_signup = YOM_TZOFE
        else:
            _v_signup = STORE  # unrecognized event text -- flagged via the note below either way
        _child_key = _participant.lower()
        if _v_signup == SEMINAR:
            _v_email = child_name_to_seminar_family_email.get(_child_key, "")
        elif _v_signup == ANNUAL:
            _v_email = child_name_to_family_email.get(_child_key, "")
        elif _v_signup == YOM_TZOFE:
            _v_email = child_name_to_yomtzofe_family_email.get(_child_key, "")
        else:
            _v_email = ""
        _v_txn_id = f"VENMO-{len(_venmo_parsed) + 1}"
        _note_bits = []
        if _v_signup == STORE:
            if "annual" not in _event_key and "seminar" not in _event_key and "yom tzofe" not in _event_key and "yomtzofe" not in _event_key:
                _note_bits.append(f"Unrecognized event text (\"{_event_raw}\") on the Venmo signups sheet -- please "
                                   f"confirm this is Annual, Seminar, Yom Tzofe, or something else; treated as Store/other for now.")
            # STORE rows are never roster-matched (same as Tapuach Store purchases from Authorize.net), so an empty
            # _v_email here doesn't mean a failed lookup -- don't also claim "no roster match" below.
        elif not _v_email:
            _note_bits.append(f"No roster match found for participant name \"{_participant}\" -- check the spelling "
                                 f"on the Venmo signups sheet against the Roster / Seminar Roster tab.")
        if _note_bits:
            venmo_notes[_v_txn_id] = " ".join(_note_bits)
        _v_row = (_v_dt.strftime("%Y-%m-%d"), _v_dt.strftime("%I:%M %p"), _v_signup, _v_amount,
                  _payee or _participant, _v_email, "Venmo", "", _v_txn_id)
        _venmo_parsed.append((_v_dt, _v_row))
    print("venmo signups loaded:", len(_venmo_parsed), "| unresolved participant name:",
          sum(1 for _d, _r in _venmo_parsed if not _r[5] and _r[2] != STORE))

# =========================================================================
# New merchandise store orders -- the "Shevet Tapuach" storefront Michal built on the Enrollia platform (sender
# info@enrollia.graysbow.com), added Sep 4, 2026, per Michal. Started operating Sep 2-4, 2026; paid via Venmo for
# now (may move to Authorize.net later). Unlike Authorize.net's Tapuach Store charges, these order-confirmation
# emails carry item-level detail (item name, size/color) and, when the order was placed on behalf of a specific
# participant (e.g. a camp t-shirt sized for a particular kid), a "For: <Participant Name>" field plus a
# "Grade: <N> Grade" field (2nd-12th grade, added Sep 4, 2026 per Michal, once she pointed out the order email
# carries it right alongside "For:") -- both captured here and shown on the Store Purchases tab via store_item_info
# below. The participant name is matched against the Annual Roster by reversed ("First Last") name since the
# roster's own Contact Name column is "Last, First"; Grade is taken as-is from the order email itself rather than
# re-derived from the Roster, since a store order can be placed for a participant not yet (or no longer) on either
# roster.
#
# Staged into data/store_orders_*.csv each refresh -- a new file per pull, same never-edit-old-files pattern as
# thread_*.csv (see STEP 4c in the daily refresh prompt) -- deduped here by Gmail message ID (never by content:
# two genuinely different orders can easily share the same item/amount, e.g. two $1.50 Hat orders).
#
# On the exact subject line to poll for: as of Sep 4, 2026 real, completed (Payment Status: Paid) orders have been
# seen under TWO different subject patterns from this same sender -- an earlier numbered one ("Order Confirmation
# #N - Shevet Tapuach", seen for orders #9-#11, all within the same half hour on Sep 3) and a later unnumbered one
# ("Order Confirmed - Shevet Tapuach", seen for two more real orders roughly a day afterward). Since the unnumbered
# format is the one seen MOST RECENTLY -- i.e. it looks like the storefront's own template changed over time rather
# than these being two structurally different kinds of order (an "admin copy" vs a "customer copy") -- STEP 4c polls
# BOTH subject patterns rather than just the numbered one, so a future template change (or reversion) doesn't
# silently stop new orders from being picked up. Both are genuine paid orders and neither has shown any sign of
# being a duplicate/second copy of the other (checked by timestamp -- no two emails share a timestamp within the
# same minute). If a future run ever finds the SAME order under both a numbered and an unnumbered subject (e.g. an
# exact date+item+amount+participant match within a minute or two of each other), don't just dedupe automatically --
# flag it in STEP 7's delivery message, since silently guessing which copy to drop risks either double-counting or
# undercounting real revenue.
_STORE_ORDERS_GLOB = os.path.join(DATA_DIR, "store_orders_*.csv")
store_item_info = {}   # STORE-order txn_id -> (item display text, participant display text, grade display text), Store Purchases tab only
store_order_notes = {}  # STORE-order txn_id -> note, merged into txn_notes below (e.g. participant not on roster)

# The storefront also sells a "$X Donation" line item (Michal, Sep 4, 2026) -- a way for a family to add a
# donation at checkout alongside (or instead of) merchandise. Detected below by item name (e.g. "$5 Donation",
# "$25 donation") and routed onto the Donations tab exactly like the auto-detected embedded-Seminar donations
# further down (merged into donations_map once that dict exists) -- it still shows on Store Purchases too (that's
# genuinely where the money came through), just with a note explaining it's also broken out on Donations.
_STORE_DONATION_RE = re.compile(r"^\$?\s*(\d+(?:\.\d{1,2})?)\s*donation\b", re.IGNORECASE)
store_order_donations = {}  # STORE-order txn_id -> donation amount, merged into donations_map below

_annual_first_last_lookup = {}  # "first last" (lower) -> canonical roster Contact Name ("Last, First"), Annual roster
for _rr in roster_rows:
    _cn = _rr["Contact Name"].strip()
    if "," in _cn:
        _last, _first = [p.strip() for p in _cn.split(",", 1)]
        if _first and _last:
            _annual_first_last_lookup[f"{_first} {_last}".lower()] = _cn
_seminar_first_last_lookup = {}  # same, but against the Seminar Roster's "Participant Name" column
for _srow in seminar_roster_rows:
    _pn = _srow["Participant Name"].strip()
    if "," in _pn:
        _last, _first = [p.strip() for p in _pn.split(",", 1)]
        if _first and _last:
            _seminar_first_last_lookup[f"{_first} {_last}".lower()] = _pn

_raw_store_orders = []
for _path in sorted(glob.glob(_STORE_ORDERS_GLOB)):
    _raw_store_orders += _load_tx_csv(_path)

_seen_store_msg_id = set()
_store_order_parsed = []
for _sr in _raw_store_orders:
    _msg_id = (_sr.get("gmail_msg_id") or "").strip()
    if not _msg_id or _msg_id in _seen_store_msg_id:
        continue  # no message ID (malformed row) or an exact re-pull of an already-captured email -- skip either way
    _seen_store_msg_id.add(_msg_id)
    _s_date_raw = (_sr.get("date") or "").strip()
    try:
        _s_dt = _dt.strptime(_s_date_raw, "%Y-%m-%d")
    except ValueError:
        _s_dt = BUILD_DATE  # unparseable/blank date -- fall back to today rather than dropping the row
    try:
        _s_amount = float((_sr.get("amount") or "0").replace("$", "").replace(",", ""))
    except ValueError:
        _s_amount = 0.0
    _buyer_name = (_sr.get("buyer_name") or "").strip()
    _buyer_email = (_sr.get("buyer_email") or "").strip().lower()
    _item = (_sr.get("item") or "").strip()
    _size_color = (_sr.get("size_color") or "").strip()
    _participant = (_sr.get("participant") or "").strip()
    _grade = (_sr.get("grade") or "").strip()
    _order_ref = (_sr.get("order_ref") or "").strip()
    _txn_id = f"STOREORDER-{_msg_id}"
    _item_display = _item + (f" ({_size_color})" if _size_color else "")
    store_item_info[_txn_id] = (_item_display, _participant, _grade)
    _donation_match = _STORE_DONATION_RE.match(_item)
    if _donation_match:
        store_order_donations[_txn_id] = float(_donation_match.group(1))
    if _participant:
        _pkey = _participant.strip().lower()
        if _pkey not in _annual_first_last_lookup and _pkey not in _seminar_first_last_lookup:
            store_order_notes[_txn_id] = (f"Store order \"For: {_participant}\" doesn't match anyone on the Annual or "
                                           f"Seminar Roster by name -- could be a typo, or a participant not yet on either "
                                           f"roster; worth a quick check.")
    _invoice_display = f"Order #{_order_ref}" if _order_ref else "Order (unnumbered)"
    _s_row = (_s_dt.strftime("%Y-%m-%d"), _s_dt.strftime("%I:%M %p"), STORE, _s_amount,
              _buyer_name, _buyer_email, "Venmo", _invoice_display, _txn_id)
    _store_order_parsed.append((_s_dt, _s_row))
if _store_order_parsed:
    print("new-storefront orders loaded:", len(_store_order_parsed), "| unmatched participant name:",
          sum(1 for _txn_id in store_order_notes))

_parsed_tx = sorted(_parsed_tx + _venmo_parsed + _store_order_parsed, key=lambda x: x[0])

transactions = [row for _, row in _parsed_tx]
print("transactions loaded from CSV:", len(transactions))

# Human-readable date range spanning every loaded transaction, for the Transactions tab header -- computed live so
# it stays correct as more thread_*.csv files get added on each daily refresh, instead of a hardcoded range.
if _parsed_tx:
    _tx_first, _tx_last = _parsed_tx[0][0], _parsed_tx[-1][0]
    if _tx_first.month == _tx_last.month and _tx_first.year == _tx_last.year:
        TX_DATE_RANGE = f"{_tx_first:%b %d}-{_tx_last.day}, {_tx_last.year}"
    else:
        TX_DATE_RANGE = f"{_tx_first:%b %d, %Y} - {_tx_last:%b %d, %Y}"
else:
    TX_DATE_RANGE = "no transactions yet"

# Build email -> matched roster rows lookup
from collections import defaultdict
email_to_kids = defaultdict(list)
for rr in roster_rows:
    for email_col in ("Parent 1 Email", "Parent 2 Email", "Email"):
        e = rr[email_col].strip().lower()
        if e:
            email_to_kids[e].append(rr["Contact Name"] + " (" + rr["Grade This Year"] + ")")

# Fallback for payments where the card's billing email doesn't match anything on file (e.g. a work email instead
# of the address the family registered with) -- match by the paying parent's first+last name instead. Keyed by
# parent name -> that family's own Parent 1 (or Parent 2) email, so a name-matched payment still resolves to the
# same canonical family key used everywhere else (sibling totals, Seminar fee lookup, etc).
name_to_family_key = {}
for rr in roster_rows:
    fam_email = rr["Parent 1 Email"].strip().lower() or rr["Parent 2 Email"].strip().lower() or rr["Email"].strip().lower()
    if not fam_email:
        continue
    for name_col in ("Parent 1 Contact Name", "Parent 2 Contact Name"):
        nm = rr[name_col].strip().lower()
        if nm:
            name_to_family_key.setdefault(nm, fam_email)

name_to_seminar_family_key = {}
for srow in seminar_roster_rows:
    fam_email = (srow["Parent 1 Email"].strip().lower().replace("\\_", "_")
                 or srow["Parent 2 Email"].strip().lower().replace("\\_", "_"))
    if not fam_email:
        continue
    for name_col in ("Parent 1 Contact Name", "Parent 2 Contact Name"):
        nm = srow[name_col].strip().lower()
        if nm:
            name_to_seminar_family_key.setdefault(nm, fam_email)

name_to_yomtzofe_family_key = {}
for yrow in yomtzofe_roster_rows:
    fam_email = (yrow["Parent 1 Email"].strip().lower().replace("\\_", "_")
                 or yrow["Email"].strip().lower().replace("\\_", "_"))
    if not fam_email:
        continue
    nm = yrow["Parent 1 Contact Name"].strip().lower()
    if nm:
        name_to_yomtzofe_family_key.setdefault(nm, fam_email)

def dedupe(seq):
    seen = set(); out = []
    for x in seq:
        if x not in seen:
            seen.add(x); out.append(x)
    return out

# =========================================================================
# Withdrawals, part 1 -- added Oct 6, 2026, per Michal. She maintains a small helper sheet, "Withdrawals for
# Payments tracker" (Date, parent name, participant name, event, refund type [Refund/Adjust/Both], updated
# installment amount [Adjust], refund amount [Refund]), logged whenever a family withdraws a participant
# mid-season. Only parsed far enough here (parent name + event -> signup) to know WHICH (name, signup) pairs
# have withdrawn -- used right below in the Transactions loop. The fuller detail (Authorize.net credit
# cross-check, Signups tab Status override, the new Withdrawals tab) is built after tx_rows exists -- see
# "Withdrawals, part 2" after the Scholarships section below.
#
# Per Michal, Oct 6, 2026: she removes a withdrawn participant from the live Roster/Seminar Roster/Yom Tzofe
# Roster sheet herself (not something build.py does) -- the existing sibling-tier pricing (annual_family_total/
# seminar_family_fee/yomtzofe_family_fee) already recomputes automatically off however many kids are left on
# that live sheet, so Expected Total needs no separate override here. One side effect of that approach: the
# Transactions tab's roster-match check is recomputed live against the CURRENT roster every run, so if an
# ENTIRE family withdraws and Michal removes them completely, their PAST (already-fine) transactions would
# otherwise flip to a false "No roster match found" simply because they're no longer on the roster at all. The
# set built below is consulted in the Transactions loop to catch that specific case and substitute a
# "withdrawn" note instead of a false flag.
_WITHDRAWALS_PATH = os.path.join(DATA_DIR, "withdrawals_refresh.csv")
withdrawal_rows_raw = []
if os.path.exists(_WITHDRAWALS_PATH):
    with open(_WITHDRAWALS_PATH, newline="") as f:
        withdrawal_rows_raw = list(csv.DictReader(f))

def _withdrawal_event_to_signup(event_raw):
    event_raw = (event_raw or "").strip().lower()
    if "yom tzofe" in event_raw or "yomtzofe" in event_raw:
        return YOM_TZOFE
    elif "seminar" in event_raw:
        return SEMINAR
    return ANNUAL  # "yearly"/"annual"/anything else -- the two explicit events are checked first

withdrawn_name_signup = set()  # (parent name lower, signup) -- consulted in the Transactions loop right below
for _wr in withdrawal_rows_raw:
    _w_parent = (_wr.get("parent name") or "").strip()
    _w_rtype = (_wr.get("refund type") or "").strip()
    if not _w_parent or not _w_rtype:
        continue  # blank/incomplete row -- skip rather than guess
    withdrawn_name_signup.add((_w_parent.lower(), _withdrawal_event_to_signup(_wr.get("event"))))

notes_map = {
    "kroitero94@gmail.com|" + ANNUAL: "$1.00 looks like a card-verification/test charge ahead of a recurring plan, not a real payment.",
    "k.mijelshon@gmail.com|" + ANNUAL: "2 kids on file (Yuval Elena, Gal Irma) -- confirm whether this payment covers one or both.",
    "meitalyitzhak@gmail.com|" + STORE: ("Tapuach Store merchandise purchase ($30.00) -- unrelated to Annual/Seminar signup fees. Logged "
                                          "here for a complete ledger, but excluded from the Signups tab and the Dashboard's "
                                          "by-signup-type totals."),
}

# Per-transaction notes, keyed by Transaction ID -- for one-off cases where the same note shouldn't repeat on
# every row for that payer/signup (e.g. one specific duplicate charge and its matching refund).
txn_notes = {
    "81775016104": ("Description was \"Offline Registration for Event: ... by: Michal Sheves\" -- Michal keyed this payment in "
                     "manually rather than the family paying online themselves. Matched as a Seminar signup."),
    "121779727125": ("This looks like a duplicate/double-submitted charge (same amount, same second, as the row above) -- "
                     "see the Aug 19 credit for the same invoice below, which refunds it."),
    "121784618133": "Refund (Credit) of the duplicate Aug 17 charge above -- nets that duplicate back out. Entered as a negative amount.",
}
txn_notes.update(venmo_notes)  # per-row notes for Venmo signups with an unrecognized event or unmatched participant name
txn_notes.update(store_order_notes)  # per-row notes for new-storefront orders whose "For:" participant isn't on either roster

tx_rows = []
resolved_email = {}  # txn_id -> the roster email key actually used to match (may differ from the payer's own
                      # email via the name fallback below) -- used downstream for family totals/notes.
for (date, time, signup, amount, name, email, card, invoice, txn_id) in transactions:
    alias_note = ""
    if signup == STORE:
        matched = ""
        flag = "OK"
        resolved_email[txn_id] = email.lower()
    else:
        if signup == SEMINAR:
            lookup = email_to_seminar_kids
        elif signup == YOM_TZOFE:
            lookup = email_to_yomtzofe_kids
        else:
            lookup = email_to_kids
        lookup_key = email.lower()
        kids = dedupe(lookup.get(lookup_key, []))
        if not kids:
            # No roster row has this billing email on file -- fall back to matching the payer's name against a
            # parent name on the roster (e.g. Sharon Blumenkranz paying from a work email, not the one on file).
            if signup == SEMINAR:
                alias_map = name_to_seminar_family_key
            elif signup == YOM_TZOFE:
                alias_map = name_to_yomtzofe_family_key
            else:
                alias_map = name_to_family_key
            alias_email = alias_map.get(name.strip().lower())
            if alias_email:
                alias_kids = dedupe(lookup.get(alias_email, []))
                if alias_kids:
                    kids = alias_kids
                    lookup_key = alias_email
                    alias_note = (f"Payer email ({email}) doesn't match anything on the roster -- matched by payer "
                                  f"name to the family on file instead (roster email {alias_email}).")
        resolved_email[txn_id] = lookup_key
        matched = "; ".join(kids)
        # Matching more than one kid isn't flagged as a review item (changed Aug 28, 2026, per Michal): both
        # Annual and Seminar totals are always tracked per FAMILY, not per kid, so this workbook's math never
        # actually depends on knowing which specific kid a given payment covers -- "Matched Child(ren)" still shows
        # everyone the fee is based on, it just isn't treated as something to review/reconcile. Only a payer who
        # can't be matched to anyone on the roster at all is a real problem worth flagging.
        if kids:
            flag = "OK"
        elif (name.strip().lower(), signup) in withdrawn_name_signup:
            # Added Oct 6, 2026 -- this family is logged on the Withdrawals helper sheet for this signup, so a
            # missing roster match here most likely means Michal already removed them from the live roster
            # sheet entirely, not a real data problem. See "Withdrawals, part 1" above.
            flag = "OK"
            alias_note = alias_note or (f"{name.strip()} withdrew from this signup and was removed from the live "
                                         f"roster -- no roster match is expected anymore; see the Withdrawals tab.")
        else:
            flag = "No roster match found"
    note = txn_notes.get(txn_id) or notes_map.get(email.lower() + "|" + signup, "") or alias_note
    tx_rows.append([date, time, signup, amount, name, email, card, invoice, txn_id, matched, flag, note])

# Fixed at 200 through Sep 12, 2026 -- caught Sep 13, 2026, per Michal: every COUNTIFS/SUMIFS/SUM formula on the
# Signups and Dashboard tabs that reads from the Transactions tab used a hardcoded "$200" range end. That was
# never actually tied to the real transaction count, and by Sep 2026 the season had grown past 200 real rows --
# so any transaction landing at row 201+ (e.g. a family's 2nd/3rd installment payment, added near the bottom of
# a freshly-appended thread_*.csv) was silently excluded from Signups' "# Payments"/"Total Paid" columns and every
# Dashboard total, even though it was correctly sitting right there on the Transactions tab itself. TX_LAST_ROW is
# now computed from the actual row count every time this script runs, so it always covers every real row no matter
# how large the season's transaction count grows -- no more manual bumping of a magic number.
TX_LAST_ROW = 2 + len(tx_rows) + 20  # +20 matches the shaded "add new rows here" buffer below the data (Transactions tab)

# Auto-detect and log Donation Add-ons for the embedded-donation Seminar families found above
# (seminar_family_donation). The donation is baked into the family's agreed total from the moment they signed up
# -- it isn't something that only becomes real once enough cash has come in -- so it's logged as soon as there's
# any payment for that signup, on the most recent one so far, regardless of whether the family still has more
# installments left to pay. (Corrected Aug 28, 2026, per Michal: an earlier version of this logic waited for
# total payments to cross the base fee before logging anything, which she didn't ask for and doesn't want --
# Brutman/Yonatan should show the $25 donation now, even though only 1 of 3 installments has come in.)
_seminar_pay_by_family = defaultdict(list)  # lookup key -> [(datetime, txn_id, amount), ...]
for _d, _row in _parsed_tx:
    _txn_id = _row[8]
    if _row[2] != SEMINAR:
        continue
    _lookup_key = resolved_email.get(_txn_id, _row[5])
    if _lookup_key in seminar_family_donation:
        _seminar_pay_by_family[_lookup_key].append((_d, _txn_id, _row[3]))

donations_map = {}     # txn_id -> donation amount, auto-populated onto the Transactions "Donation Add-on" column
donation_notes = {}    # txn_id -> explanatory note, prepended to that row's Notes
for _key, _pays in _seminar_pay_by_family.items():
    _donation_amt, _base_total, _raw_total = seminar_family_donation[_key]
    _pays.sort(key=lambda p: p[0])
    _running = sum(_amt for _d_unused, _txn_id_unused, _amt in _pays)
    _latest_txn_id = _pays[-1][1]
    donations_map[_latest_txn_id] = _donation_amt
    _progress = ("fully paid" if _running >= _base_total
                 else f"${_running:.2f} of the ${_base_total:.2f} base paid so far, still mid-installment-plan")
    donation_notes[_latest_txn_id] = (
        f"Auto-detected Donation Add-on: this family's Seminar Roster Participant Fee (${_raw_total:.2f}) bakes in "
        f"a ${_donation_amt:.0f} voluntary donation on top of the real ${_base_total:.2f} fee (Price Reference "
        f"only lists $345/$305 per participant). Logged here regardless of remaining installments ({_progress}).")

# Merge in the storefront's "$X Donation" line items detected above (Michal, Sep 4, 2026) -- same mechanism as the
# embedded-Seminar donations just above: populates the Transactions "Donation Add-on" column and flows straight
# onto the Donations tab.
for _store_txn_id, _store_donation_amt in store_order_donations.items():
    donations_map[_store_txn_id] = _store_donation_amt
    donation_notes[_store_txn_id] = (
        f"Storefront \"Donation\" line item (${_store_donation_amt:.0f}) -- also broken out on the Donations tab.")

for _tr in tx_rows:
    _dnote = donation_notes.get(_tr[8])
    if _dnote:
        _tr[11] = (_dnote + " " + _tr[11]) if _tr[11] else _dnote

ws = wb.create_sheet("Transactions")
ws["A1"] = (f"Raw payment ledger -- one row per Authorize.net receipt email or Venmo signup payment, {TX_DATE_RANGE} "
            f"({len(tx_rows)} transactions). Add new rows below as new receipts arrive. Column G shows \"Venmo\" for rows sourced "
            "from Michal's \"Venmo signups for tracker\" sheet (added Aug 31, 2026) instead of an Authorize.net receipt -- those "
            "have no card info and no installment plans, always paid in full. Column M auto-fills for Seminar families whose "
            "Roster fee bakes in a donation (see Read Me); for anything else, if a payment is higher than the signup price, check "
            "whether it includes an optional Donation Add-on (see Price Reference) and log it in column M yourself. Columns O/P "
            "(added Aug 30, 2026, per Michal) show Authorize.net's ~2% processing fee and what actually lands in the account after "
            "it -- fee is 0 on refunds/credits (negative amounts) and on Venmo rows, since neither is charged an Authorize.net fee.")
ws["A1"].font = SUB_FONT
ws.merge_cells("A1:P1")
ws.row_dimensions[1].height = 52

# Authorize.net's processing fee -- applied to the gross charge (column D), not just the signup portion, since
# that's what Authorize.net actually assesses the fee against. Michal, Aug 30, 2026: ~2%, and NOT charged on
# refunds/credits -- those already show as a negative Amount, so the IF below zeroes the fee out for them rather
# than computing a (nonsensical) negative fee. Also zeroed for Venmo rows (column G = "Venmo", added Aug 31,
# 2026) -- Venmo signup payments never go through Authorize.net, so no such fee applies to them at all.
AUTHNET_FEE_RATE = 0.02

headers = ["Date", "Time", "Signup", "Amount", "Payer Name", "Payer Email", "Card / Payment Method", "Invoice #",
           "Transaction ID", "Matched Child(ren)", "Review Flag", "Notes", "Donation Add-on\n(enter manually)",
           "Net Toward Signup Fee", "Authorize.net Fee\n(2%, none on refunds/Venmo)", "Amount After Fee"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

donation_dv = DataValidation(type="list", formula1='"25,50,100,200,500,1000"', allow_blank=True,
                              showDropDown=False, error="Pick a donation amount or leave blank.",
                              errorTitle="Donation Add-on")
ws.add_data_validation(donation_dv)

for i, row in enumerate(tx_rows):
    rownum = 3 + i
    body_row(ws, rownum, row, len(row), start_col=1)
    ws.cell(row=rownum, column=4).number_format = CURRENCY
    flagcell = ws.cell(row=rownum, column=11)
    if flagcell.value == "OK":
        flagcell.fill = GREEN_FILL
    else:
        flagcell.fill = ORANGE_FILL
    don_cell = ws.cell(row=rownum, column=13, value=donations_map.get(row[8]))
    don_cell.fill = YELLOW_FILL
    don_cell.font = BODY_FONT
    don_cell.border = BORDER
    don_cell.number_format = CURRENCY
    donation_dv.add(don_cell)
    net_cell = ws.cell(row=rownum, column=14,
                        value=f'=D{rownum}-IF(M{rownum}="",0,M{rownum})')
    net_cell.font = BODY_FONT
    net_cell.border = BORDER
    net_cell.number_format = CURRENCY
    fee_cell = ws.cell(row=rownum, column=15,
                        value=f'=IF(AND(D{rownum}>0,G{rownum}<>"Venmo"),D{rownum}*{AUTHNET_FEE_RATE},0)')
    fee_cell.font = BODY_FONT
    fee_cell.border = BORDER
    fee_cell.number_format = CURRENCY
    afterfee_cell = ws.cell(row=rownum, column=16, value=f'=D{rownum}-O{rownum}')
    afterfee_cell.font = BODY_FONT
    afterfee_cell.border = BORDER
    afterfee_cell.number_format = CURRENCY

# Shaded "add new rows here" area
for r in range(3+len(tx_rows), 3+len(tx_rows)+20):
    for c in range(1, 17):
        ws.cell(row=r, column=c).fill = GREY_FILL
        ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=14, value=f'=D{r}-IF(M{r}="",0,M{r})').number_format = CURRENCY
    ws.cell(row=r, column=15, value=f'=IF(AND(D{r}>0,G{r}<>"Venmo"),D{r}*{AUTHNET_FEE_RATE},0)').number_format = CURRENCY
    ws.cell(row=r, column=16, value=f'=D{r}-O{r}').number_format = CURRENCY
    donation_dv.add(ws.cell(row=r, column=13))

ws.freeze_panes = "A3"
autofit(ws, [11, 9, 40, 11, 20, 26, 18, 22, 15, 30, 26, 46, 15, 15, 16, 16])
ws.auto_filter.ref = f"A2:P{2+len(tx_rows)}"

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("transactions written:", len(tx_rows))
for row in tx_rows:
    print(row[4], "|", row[2][:20], "|", row[9], "|", row[10])
# Added Sep 20, 2026, per Michal -- a real Authorize.net payment (Ronit Grofman) sat on the Transactions tab all
# morning flagged "No roster match found" and nothing in this script's own printed summary ever said so; Michal
# only caught it by independently cross-checking the source Roster sheet herself. Venmo (STEP 4b) and the
# storefront (STEP 4c) already get an "unresolved/unmatched participant name" count in the summary -- Authorize.net
# receipts (STEP 4, the largest and most common source) never did. This closes that gap so an unmatched payer name
# is impossible to miss in STEP 5/STEP 7 of the daily run, the same as the other two sources already are.
_unmatched_roster_rows = [row for row in tx_rows if row[10] == "No roster match found"]
print("Annual/Seminar transactions with no roster match:", len(_unmatched_roster_rows),
      ("| " + "; ".join(f"{row[4]} ({row[2][:20].strip()}, {row[8]})" for row in _unmatched_roster_rows) if _unmatched_roster_rows else ""))

# =========================================================================
# SHEET: Reconciliation (Aug 29, 2026, per Michal) -- cross-checks the receipt-email-derived Transactions tab
# against Authorize.net's own "Transaction Detail" export, which Michal pulls manually from the Authorize.net
# merchant portal every couple of weeks (not daily) into a separate Google Sheet titled "Authorize transaction
# report" (source: https://docs.google.com/spreadsheets/d/1j2KFTe0iYJ6kgw0UzYjLuomS1iRrpDmKkZ7RXq1evU8). Since it's
# refreshed on Michal's own schedule rather than automatically, data/authorize_report.csv is treated as optional --
# if it isn't present (or is stale relative to today), this tab still renders with a note rather than erroring out.
# Purpose: catch any real payment whose receipt email was missed by the Gmail search/parsing in the Transactions
# tab above (e.g. buried in a long thread, or a "Summary of Automated Recurring Billing" digest that never got its
# own line-item receipt email -- see Read Me's Known Limitations). A transaction that's in the Authorize.net report
# but NOT in Transactions is a real gap worth investigating; the reverse (in Transactions but not yet in this
# report) is expected and not flagged, since Transactions is refreshed daily and this report lags behind it.
# =========================================================================
_AUTH_REPORT_PATH = os.path.join(DATA_DIR, "authorize_report.csv")

def _parse_authnet_dt(s):
    # e.g. "10-Aug-2026 03:50:51 AM EDT" -> strip trailing tz abbreviation, parse 12-hour clock.
    s = re.sub(r"\s+[A-Z]{2,4}$", "", s.strip())
    return _dt.strptime(s, "%d-%b-%Y %I:%M:%S %p")

def _settled_amount(rr):
    # Added Oct 1, 2026, per Michal -- her Sep refresh of the "Authorize transaction report" sheet came through
    # without a "Settlement Amount" column at all (it also dropped "Transaction Status", used only for display
    # below). That column had real values when the Payouts tab's Cumulative Gap logic was tested against it on
    # Sep 4, 2026 (see the Payouts tab comments -- that test's $90,930.41 "our settled total" came from this exact
    # field), so this isn't a brand-new design gap, just this particular Authorize.net portal export not including
    # it this time (maybe a different column selection in the portal, or a report-template change on their end).
    # Falling back silently to 0 whenever the column is simply missing (the prior behavior) would make every date's
    # "our settled total" read as $0 and the Cumulative Gap balloon to ~100% of batch volume -- not a real gap,
    # just a missing column -- which is exactly the kind of misleading-but-silent failure this pipeline tries hard
    # to avoid elsewhere (see the Reconciliation/Annual-Seminar-unmatched notes above). So: use the real Settlement
    # Amount when the column is present and populated; otherwise fall back to a reasoned proxy from Total Amount +
    # Action Code + Response Code (Authorize.net's own approved/declined/error flag) -- Response Code "1" means
    # approved, so only approved rows ever contribute: AUTH_CAPTURE settles at +Total Amount, CREDIT settles at
    # -Total Amount (a refund), VOID never settles (never charged) so contributes $0, and a declined/errored
    # attempt (Response Code 2/3) also contributes $0 even though it still gets a Settlement Date/Time stamped on
    # it in this export (confirmed against this month's data -- 11 declined + 3 error rows all have one).
    _raw = (rr.get("Settlement Amount") or "").strip()
    if _raw:
        try:
            return float(_raw)
        except ValueError:
            pass
    if rr.get("Response Code", "").strip() != "1":
        return 0.0
    try:
        _total = float(rr.get("Total Amount") or 0)
    except ValueError:
        return 0.0
    _action = rr.get("Action Code", "").strip()
    if _action == "AUTH_CAPTURE":
        return _total
    if _action == "CREDIT":
        return -_total
    return 0.0  # VOID, or any other/unrecognized Action Code -- never settled, don't guess

recon_rows = []       # rows to display: (Transaction ID, Submit Date, Name, Email, Amount, Invoice Desc, Status)
recon_report_count = 0
recon_report_range = ""
recon_available = os.path.exists(_AUTH_REPORT_PATH)

if recon_available:
    with open(_AUTH_REPORT_PATH, newline="") as f:
        _report_rows = list(csv.DictReader(f))
    recon_report_count = len(_report_rows)

    _report_parsed = []
    for _rr in _report_rows:
        _tid = _rr.get("Transaction ID", "").strip()
        if not _tid:
            continue
        try:
            _rdt = _parse_authnet_dt(_rr.get("Submit Date/Time", ""))
        except ValueError:
            _rdt = None
        _report_parsed.append((_tid, _rdt, _rr))

    _dated = [p for p in _report_parsed if p[1] is not None]
    if _dated:
        _dated.sort(key=lambda p: p[1])
        _r_first, _r_last = _dated[0][1], _dated[-1][1]
        recon_report_range = f"{_r_first:%b %d, %Y} - {_r_last:%b %d, %Y}"

    for _tid, _rdt, _rr in _report_parsed:
        if _tid in _seen_txnid:
            continue  # already accounted for via a receipt email -- no gap
        if _rr.get("Response Code", "").strip() != "1":
            # Added Oct 5, 2026, per Michal -- a run flagged 4 "gaps" (Orital Karelic $2,130 Aug 23, Gili Sharett
            # $183 Aug 25, Avishag Iluz $740 Aug 27, Amy Nachon $1,550 Sep 27) that a manual Gmail search couldn't
            # explain. Investigation found all 4 (and in fact every single gap in the report at the time, 14 of 14)
            # share one thing: Response Code "2" (declined) or "3" (error) in Authorize.net's own report -- a
            # declined/errored authorization attempt never gets a receipt email at all (nothing was actually
            # charged), so it correctly never appears in data/thread_*.csv. In 3 of those 4 cases the same family's
            # very next attempt for the same amount has Response Code "1" and DOES have a matching receipt -- a
            # normal declined-then-retried-successfully pattern, not a missing receipt. This loop previously had no
            # Response Code check at all, so every declined/errored attempt that Authorize.net happens to log was
            # flagged as a false gap. Same Response Code "1" == approved convention _settled_amount() above already
            # relies on -- skip anything that isn't a genuinely approved transaction before treating it as a gap.
            continue
        _name = (_rr.get("Customer First Name", "").strip() + " " + _rr.get("Customer Last Name", "").strip()).strip()
        recon_rows.append([
            _tid,
            _rdt.strftime("%Y-%m-%d %I:%M %p") if _rdt else _rr.get("Submit Date/Time", ""),
            _name,
            _rr.get("Email", "").strip(),
            float(_rr.get("Total Amount") or 0),
            _rr.get("Invoice Description", "").strip(),
            _rr.get("Transaction Status", "").strip(),
        ])
    recon_rows.sort(key=lambda r: r[1])

ws = wb.create_sheet("Reconciliation")
if recon_available:
    _summary = (f"Checked {recon_report_count} transactions from Authorize.net's own report ({recon_report_range}) against the "
                f"{len(_seen_txnid)} transactions logged on the Transactions tab (all pulled from receipt emails). Declined/errored "
                f"attempts (Response Code 2/3) are excluded below -- they never get a receipt email, so they're not real gaps. ")
    if recon_rows:
        _summary += (f"{len(recon_rows)} transaction(s) below appear in Authorize.net's report but were NOT found among the parsed "
                      "receipt emails -- these are worth a manual look ( search Gmail directly for the Transaction ID / amount / name).")
    else:
        _summary += "No gaps found -- every transaction in Authorize.net's report is accounted for on the Transactions tab."
else:
    _summary = ("No Authorize.net report has been loaded yet (data/authorize_report.csv not found). Michal pulls a \"Transaction "
                "Detail\" export from the Authorize.net merchant portal every couple of weeks into a separate Google Sheet "
                "(\"Authorize transaction report\") for this cross-check; once that's refreshed and re-pulled, this tab will "
                "compare it against the receipt-email data above.")
ws["A1"] = _summary
ws["A1"].font = SUB_FONT
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A1:G1")
ws.row_dimensions[1].height = 42

headers = ["Transaction ID", "Submit Date/Time", "Customer Name", "Email", "Amount", "Invoice Description", "Authorize.net Status"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

for i, row in enumerate(recon_rows):
    rownum = 3 + i
    body_row(ws, rownum, row, len(row), start_col=1, fills={4: ORANGE_FILL})
    ws.cell(row=rownum, column=5).number_format = CURRENCY

last_row = 2 + len(recon_rows)
if not recon_rows:
    _msg = "No gaps to show." if recon_available else "Nothing to compare yet -- see note above."
    ws.cell(row=3, column=1, value=_msg).font = SUB_FONT
    last_row = 3

ws.freeze_panes = "A3"
autofit(ws, [15, 20, 22, 26, 12, 40, 20])
ws.auto_filter.ref = f"A2:G{max(last_row, 2)}"

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("reconciliation gaps found:", len(recon_rows), "| report loaded:", recon_available)

# =========================================================================
# SHEET: Payouts (added Sep 4, 2026, per Michal) -- cross-checks Payments Hub's own daily bank-deposit report against
# Authorize.net's settled-transaction data, catching anything that went missing on the way from card charge to bank
# account. Different from the Reconciliation tab above: Reconciliation checks "did every Authorize.net transaction
# get a receipt email logged on Transactions"; this tab checks "did every dollar Authorize.net settled actually show
# up as bank-side batch volume/deposit" -- a check from the bank's own report, not the gateway's.
#
# Source: "Payments Hub" (hello@paymentshub.com), the merchant acquirer that actually funds the org's bank account --
# Authorize.net is just the payment gateway in front of it, so these are two views of the same underlying settled
# batches, not two unrelated systems. Payments Hub emails a "Daily Credit/Debit Snapshot" each morning covering the
# PREVIOUS calendar day, to tapuach@israelscouts.org; Michal set up a Gmail filter (Sep 4, 2026) forwarding these to
# gizbartapuach@israelscouts.org, which is what this pipeline actually reads (see STEP 4d in the daily refresh
# prompt). Staged into data/payouts_*.csv each refresh, same never-edit-old-files/dedupe-by-gmail_msg_id pattern as
# store_orders_*.csv above.
#
# Matching logic (REVISED Sep 4, 2026, after testing against real overlapping data -- see below): "Total batch
# volume" for a given date is compared against the SUM of authorize_report.csv's Settlement Amount, grouped by that
# report's own Settlement Date/Time field (not Submit Date/Time). The original design ( see git/version history)
# treated this as a same-day, tight-tolerance match and flagged any day that didn't line up exactly as "Review."
# Once Michal forwarded her full inbox of Payments Hub emails back to the start of signups (Aug 10, 2026) and her
# Authorize.net report sheet's coverage (through Aug 29, 2026) actually overlapped real payout dates for the first
# time, that same-day hypothesis turned out to be too strict: only isolated, single-transaction days lined up
# exactly on their own calendar date -- as soon as a day had multiple transactions, the settlement date Authorize.net
# recorded routinely fell a day (sometimes more) off from the date Payments Hub's batch closed, so a same-day $1
# tolerance flagged 17 of 18 real overlapping days as "Review" even though nothing was actually wrong -- it's normal
# settlement lag, not missing money.
#
# What DOES hold up: tracked as a running CUMULATIVE total across the whole overlap window (Aug 10-29, 2026),
# cumulative Total Batch Volume ($93,604.12) and cumulative Authorize.net settled total ($90,930.41) stayed within
# a bounded, oscillating gap that shrank back down repeatedly rather than drifting further apart -- ending at a
# $2,673.71 gap (2.9% of volume) by the end of the window, which is fully explainable as the last couple of days'
# batch volume not having "caught up" into a recorded settlement date yet inside the report's own cutoff. That
# self-correcting, bounded-percentage pattern is the actual signal for "is anything major missing" -- not any single
# day's exact match. So this tab now shows the day-by-day settlement-date comparison as REFERENCE ONLY (no
# Matched/Review verdict per row -- that call turned out not to be reliable at daily grain), and instead tracks a
# running Cumulative Gap ($) and Cumulative Gap (%) column: if that running percentage stays in the low single
# digits and keeps returning there after temporary spikes, that's healthy; a gap that keeps climbing and never
# comes back down over an extended stretch would be the real red flag worth investigating.
#
# "Total deposit amount" (the net $ that actually hit the bank that calendar day) is shown for reference but
# deliberately NOT forced into a same-day match against our computed (batch total - fee): real data already shows
# deposits lag batch volume by an uneven number of days (Aug 30, 2026 shows $0 deposit against $5,420 in batch
# volume that same day -- nothing landed in the bank yet), so a naive same-day comparison would falsely flag every
# single day as a mismatch. "Month-to-date net deposits" ( also from the email, a running total) is the more
# reliable big-picture check for "is money actually showing up over time" -- shown alongside for that reason.
#
# Known gap in the source data: Michal forwarded every Payments Hub snapshot email she has going back to the start
# of signups, and the very first one in her inbox covers Aug 10, 2026 -- but there's no snapshot for Aug 11 or
# Aug 12, 2026 (the next one jumps to Aug 13). Every other day in the whole range has one, including other $0-volume
# days, so this looks like a real 2-day hole rather than "no email sent" -- worth Michal checking her tapuach inbox
# (including spam/trash) for those two dates.
# =========================================================================
_payout_files = sorted(glob.glob(os.path.join(DATA_DIR, "payouts_*.csv")))
payout_rows_raw = []
_seen_payout_msg_id = set()
for _po_path in _payout_files:
    with open(_po_path, newline="") as f:
        _por_rows = list(csv.DictReader(f))
    for _por in _por_rows:
        _msg_id = (_por.get("gmail_msg_id") or "").strip()
        if not _msg_id or _msg_id in _seen_payout_msg_id:
            continue
        _seen_payout_msg_id.add(_msg_id)
        payout_rows_raw.append(_por)
payout_rows_raw.sort(key=lambda r: r.get("date", ""))
payouts_available = len(payout_rows_raw) > 0

# Our own settled totals, grouped by Authorize.net's Settlement Date (date part only) -- used to check batch volume.
_settled_by_date = defaultdict(float)
_max_settle_date = None
if recon_available:
    for _rr in _report_rows:
        _sd_raw = (_rr.get("Settlement Date/Time") or "").strip()
        if not _sd_raw:
            continue
        try:
            _sdt = _parse_authnet_dt(_sd_raw)
        except ValueError:
            continue
        _sdate = _sdt.date()
        if _max_settle_date is None or _sdate > _max_settle_date:
            _max_settle_date = _sdate
        _samt = _settled_amount(_rr)
        _settled_by_date[_sdate] += _samt

payout_display_rows = []  # (date, deposit, batch_vol, txn_count, refunds_disp, batch_vol_net, our_total_disp, diff_disp, cum_gap, cum_gap_pct, status, mtd, notes)
_cum_batch = 0.0
_cum_settle = 0.0
_last_cum_gap = None
_last_cum_gap_pct = None
_last_cum_date = None
for _por in payout_rows_raw:
    _p_date_str = (_por.get("date") or "").strip()
    try:
        _p_date = _dt.strptime(_p_date_str, "%Y-%m-%d").date()
    except ValueError:
        _p_date = None
    try:
        _deposit = float(_por.get("total_deposit_amount") or 0)
    except ValueError:
        _deposit = 0.0
    try:
        _batch_vol = float(_por.get("total_batch_volume") or 0)
    except ValueError:
        _batch_vol = 0.0
    _txn_count = (_por.get("transactions_count") or "").strip()
    try:
        _refunds_amt = float(_por.get("refunds_amount") or 0)
    except ValueError:
        _refunds_amt = 0.0
    _refunds_count = (_por.get("refunds_count") or "0").strip()
    _refunds_disp = f"{_refunds_count} totaling ${_refunds_amt:,.2f}" if _refunds_count not in ("", "0") else ""
    try:
        _mtd = float(_por.get("mtd_net_deposits") or 0)
    except ValueError:
        _mtd = 0.0

    # Added Oct 15, 2026, per Michal ("are the payouts ... also taking into account the refunds?") -- Payments
    # Hub's own "Total Batch Volume" field is GROSS: it does not subtract that day's refunds, which the email
    # instead reports in its own separate "Refunds" field (confirmed against real data -- Oct 14, 2026 shows
    # $2,065.02 batch volume alongside $1,939.18 in refunds that same day, which is exactly the day's 4 withdrawal
    # refunds added to Transactions). Our own side of the comparison (_settled_amount/_settled_by_date, used for
    # "Our Authorize.net Total" below) is already NET -- a CREDIT (refund) row subtracts there. Comparing a gross
    # bank figure against a net Authorize.net figure would make any refund-heavy day look like a large fake
    # settlement gap, so _batch_vol_net (batch volume minus that day's reported refunds) is what actually feeds
    # the Day Difference / Cumulative Gap math below; the raw, as-reported Total Batch Volume stays in its own
    # column (C) for reference/audit, unchanged.
    _batch_vol_net = _batch_vol - _refunds_amt

    _in_range = _p_date is not None and recon_available and _max_settle_date is not None and _p_date <= _max_settle_date
    if _in_range:
        _our_total = _settled_by_date.get(_p_date, 0.0)
        _diff = _batch_vol_net - _our_total
        _our_total_disp = _our_total
        _diff_disp = _diff
        _cum_batch += _batch_vol_net
        _cum_settle += _our_total
        _cum_gap = _cum_batch - _cum_settle
        _cum_gap_pct = (_cum_gap / _cum_batch) if _cum_batch else 0.0
        _last_cum_gap, _last_cum_gap_pct, _last_cum_date = _cum_gap, _cum_gap_pct, _p_date
        _status = "Reference -- see Cumulative Gap"
    elif not recon_available:
        _our_total_disp = None
        _diff_disp = None
        _cum_gap = None
        _cum_gap_pct = None
        _status = "Pending -- no Authorize.net report loaded yet"
    else:
        _our_total_disp = None
        _diff_disp = None
        _cum_gap = None
        _cum_gap_pct = None
        _status = (f"Pending -- Authorize.net report only covers through {_max_settle_date:%b %d, %Y}" if _max_settle_date
                    else "Pending -- no Authorize.net report loaded yet")

    _notes_parts = []
    for _label, _key in (("disputes", "disputes_count"), ("new dispute inquiries", "new_disputes_inquiries"),
                           ("chargebacks", "new_disputes_chargebacks"), ("ACH rejections", "ach_rejections_count")):
        try:
            _n = int(float(_por.get(_key) or 0))
        except ValueError:
            _n = 0
        if _n:
            _notes_parts.append(f"{_n} {_label}")
    _notes = "; ".join(_notes_parts)

    payout_display_rows.append([
        _p_date.strftime("%Y-%m-%d") if _p_date else _p_date_str,
        _deposit, _batch_vol, _txn_count, _refunds_disp, _batch_vol_net, _our_total_disp, _diff_disp, _cum_gap, _cum_gap_pct, _status, _mtd, _notes,
    ])

ws = wb.create_sheet("Payouts")
if payouts_available:
    _in_range_n = sum(1 for r in payout_display_rows if r[9] == "Reference -- see Cumulative Gap")
    _pending_n = len(payout_display_rows) - _in_range_n
    if _last_cum_gap_pct is not None:
        _health = ("small and healthy" if abs(_last_cum_gap_pct) < 0.10 else
                    "worth keeping an eye on" if abs(_last_cum_gap_pct) < 0.20 else
                    "large enough to look into")
        _cum_sentence = (f"As of {_last_cum_date:%b %d, %Y} (the most recent date both sources cover), the running Cumulative Gap "
                          f"is ${_last_cum_gap:,.2f} ({_last_cum_gap_pct*100:.1f}% of cumulative batch volume, net of refunds) -- "
                          f"{_health}. Day-by-day, the settlement date Authorize.net records for a transaction routinely lands a day "
                          "or more off from the date Payments Hub's batch closed, so per-day differences (columns G/H) are shown for "
                          "reference only, not as a pass/fail -- the Cumulative Gap columns are the more reliable \"is anything major "
                          "missing\" check, since a gap that stays small and keeps shrinking back down is just normal settlement lag, "
                          "not missing money.")
    else:
        _cum_sentence = ""
    _summary = (f"{len(payout_display_rows)} daily payout snapshot(s) from Payments Hub (the bank-side processor) -- {_in_range_n} "
                f"within Authorize.net's report coverage, {_pending_n} still \"Pending\" until that report catches up. {_cum_sentence} "
                "Added Oct 15, 2026, per Michal: Payments Hub's own \"Total Batch Volume\" (column C) is gross -- it does not "
                "subtract that day's refunds, which it reports separately in \"Refunds\" (column E) -- so a new \"Batch Volume Net "
                "of Refunds\" column (F) subtracts them, and THAT net figure (not the raw gross one) is what the Day Difference / "
                "Cumulative Gap columns actually compare against Authorize.net's own already-refund-netted settled total; the raw "
                "gross Total Batch Volume stays in column C unchanged, for reference. Total Deposit Amount and Month-to-Date Net "
                "Deposits are shown for reference too, but not forced into a same-day match, since bank funding lags the batch by an "
                "uneven number of days (see Read Me).")
else:
    _summary = ("No payout data loaded yet ( no data/payouts_*.csv files found). Michal forwards Payments Hub's daily bank-deposit "
                "snapshot emails to gizbartapuach@israelscouts.org; once the daily refresh has pulled at least one, this tab will "
                "populate.")
ws["A1"] = _summary
ws["A1"].font = SUB_FONT
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A1:M1")
ws.row_dimensions[1].height = 84

headers = ["Date", "Total Deposit Amount", "Total Batch Volume (gross)", "Transactions", "Refunds",
           "Batch Volume Net of Refunds", "Our Authorize.net Total (by Settlement Date, net of refunds)",
           "Day Difference (Net Batch Volume - Ours)",
           "Cumulative Gap ($)", "Cumulative Gap (%)", "Status", "Month-to-Date Net Deposits", "Notes"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

for i, row in enumerate(payout_display_rows):
    rownum = 3 + i
    _status_fill = GREY_FILL
    vals = (row[:6]
            + [row[6] if row[6] is not None else "--", row[7] if row[7] is not None else "--",
               row[8] if row[8] is not None else "--", row[9] if row[9] is not None else "--"]
            + row[10:])
    body_row(ws, rownum, vals, len(vals), start_col=1, fills={10: _status_fill})
    for col in (2, 3, 6, 7, 8, 9, 10, 12):
        c = ws.cell(row=rownum, column=col)
        if isinstance(c.value, (int, float)):
            c.number_format = '0.0%' if col == 10 else CURRENCY

last_row = 2 + len(payout_display_rows)
if payout_display_rows:
    ws.cell(row=last_row + 2, column=1, value="Total").font = BOLD_FONT
    for col in (2, 3, 6):
        col_letter = get_column_letter(col)
        tot_cell = ws.cell(row=last_row + 2, column=col, value=f"=SUM({col_letter}3:{col_letter}{last_row})")
        tot_cell.font = BOLD_FONT
        tot_cell.number_format = CURRENCY
else:
    ws.cell(row=3, column=1, value="Nothing to show yet -- see note above.").font = SUB_FONT
    last_row = 3

ws.freeze_panes = "A3"
autofit(ws, [12, 16, 16, 12, 16, 18, 26, 20, 16, 14, 26, 18, 30])
ws.auto_filter.ref = f"A2:M{max(last_row, 2)}"

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("payout rows loaded:", len(payout_display_rows), "| in Authorize.net coverage:",
      sum(1 for r in payout_display_rows if r[10] == "Reference -- see Cumulative Gap"),
      "| final cumulative gap:", f"{_last_cum_gap:,.2f} ({_last_cum_gap_pct*100:.1f}%)" if _last_cum_gap is not None else "n/a")

# =========================================================================
# SHEET: Payout by Event (added Sep 11, 2026, per Michal) -- estimates how much of the ACTUAL bank-side payout
# (from the Payouts tab / Payments Hub) is attributable to each event (Annual Membership, Opening Seminar,
# Tapuach Store), as a CUMULATIVE running total by calendar month starting August 2026 (September's figures
# include both September and August, and so on). Michal asked for this to see, month over month, roughly how
# much of what's actually landed in the bank so far can be attributed to each event. Designed to be delivered by
# a new monthly scheduled task that fires the 2nd of each month covering the previous month.
#
# IMPORTANT SCOPE NOTE: this can only be based on Authorize.net-settled activity (card payments that flow
# Authorize.net -> Payments Hub -> bank), because that's the only data with both (a) a bank-side payout record
# and (b) a per-transaction event label -- the Invoice Description field on data/authorize_report.csv, per
# Michal's reminder (Sep 11, 2026), classified with the same ANNUAL/SEMINAR/STORE substring match used
# elsewhere in this script. Venmo-paid signups/orders never touch Payments Hub at all -- there's no bank payout
# record to allocate for them -- so Venmo-sourced revenue is NOT represented here, even though it appears on the
# Signups/Donations/Store Purchases tabs. See the note cell below and Read Me.
#
# Methodology: for each calendar month end (Aug 2026 onward), compute CUMULATIVE (since Aug 1, 2026) totals:
#   1. Cumulative settled $ per event, from authorize_report.csv's Settlement Amount, grouped by the event
#      classification of Invoice Description, restricted to Settlement Date/Time falling in or before that month.
#   2. Cumulative actual bank payout (Total Batch Volume from data/payouts_*.csv), restricted the same way the
#      Payouts tab restricts its comparison (payout date <= the Authorize.net report's own latest settlement
#      coverage date), so this never compares bank data against a period the gateway report hasn't caught up to.
#   3. An allocation ratio = cumulative batch volume / cumulative settled $ (this is the same kind of gap the
#      Payouts tab already tracks as "Cumulative Gap" -- processing fees, timing lag, etc.). Each event's
#      ESTIMATED payout = that event's cumulative settled $ x this ratio. This is an estimate, not an exact
#      reconciliation -- Payments Hub's own report has no per-event breakdown, only daily aggregate batch/deposit
#      totals, so this is the best available approximation of real bank dollars per event.
# =========================================================================
_EVENT_DISPLAY_PBE = {ANNUAL: "Annual Membership", SEMINAR: "Opening Seminar", YOM_TZOFE: "Yom Tzofe", STORE: "Tapuach Store"}
_EVENT_ORDER_PBE = [ANNUAL, SEMINAR, YOM_TZOFE, STORE]
_PBE_SCOPE_START = _dt(2026, 8, 1).date()

def _classify_desc_pbe(desc):
    d = (desc or "")
    if ANNUAL in d:
        return ANNUAL
    if SEMINAR in d:
        return SEMINAR
    if YOM_TZOFE in d:
        return YOM_TZOFE
    return STORE

# Settled $ per (event, settlement-month) from the Authorize.net report, Aug 2026 onward.
_settled_by_event_month = defaultdict(float)  # (event, (year, month)) -> $
_pbe_months = set()
if recon_available:
    for _rr in _report_rows:
        _sd_raw = (_rr.get("Settlement Date/Time") or "").strip()
        if not _sd_raw:
            continue
        try:
            _sdt = _parse_authnet_dt(_sd_raw)
        except ValueError:
            continue
        _sdate = _sdt.date()
        if _sdate < _PBE_SCOPE_START:
            continue
        _samt = _settled_amount(_rr)
        _event = _classify_desc_pbe(_rr.get("Invoice Description", ""))
        _month_key = (_sdate.year, _sdate.month)
        _settled_by_event_month[(_event, _month_key)] += _samt
        _pbe_months.add(_month_key)

# Actual bank batch volume per payout-date-month, Aug 2026 onward, restricted to dates within Authorize.net's own
# report coverage (same restriction the Payouts tab uses -- payout date <= _max_settle_date) so we don't compare
# against days the gateway report hasn't caught up to yet. Net of that day's refunds (added Oct 15, 2026, per
# Michal -- same fix as the Payouts tab above): Payments Hub's "Total Batch Volume" is gross, reporting refunds
# separately in its own "Refunds" field, while the settled $ this gets compared against (_settled_by_event_month,
# via _settled_amount) is already net of credits -- so this subtracts refunds_amount here too, for the same
# apples-to-apples reason, before it ever reaches the allocation-ratio math below.
_batch_by_month = defaultdict(float)
for _por in payout_rows_raw:
    _p_date_str = (_por.get("date") or "").strip()
    try:
        _p_date = _dt.strptime(_p_date_str, "%Y-%m-%d").date()
    except ValueError:
        continue
    if _p_date < _PBE_SCOPE_START:
        continue
    if not (recon_available and _max_settle_date is not None and _p_date <= _max_settle_date):
        continue
    try:
        _batch_vol = float(_por.get("total_batch_volume") or 0)
    except ValueError:
        _batch_vol = 0.0
    try:
        _refunds_amt_pbe = float(_por.get("refunds_amount") or 0)
    except ValueError:
        _refunds_amt_pbe = 0.0
    _month_key = (_p_date.year, _p_date.month)
    _batch_by_month[_month_key] += (_batch_vol - _refunds_amt_pbe)
    _pbe_months.add(_month_key)

_pbe_sorted_months = sorted(_pbe_months)

pbe_rows = []  # one row per month: cumulative figures as of that month's end (or latest data available)
_cum_settled_by_event_pbe = defaultdict(float)
_cum_batch_pbe = 0.0
for _mk in _pbe_sorted_months:
    for _event in _EVENT_ORDER_PBE:
        _cum_settled_by_event_pbe[_event] += _settled_by_event_month.get((_event, _mk), 0.0)
    _cum_batch_pbe += _batch_by_month.get(_mk, 0.0)
    _cum_settled_total_pbe = sum(_cum_settled_by_event_pbe.values())
    _ratio_pbe = (_cum_batch_pbe / _cum_settled_total_pbe) if _cum_settled_total_pbe else None
    _est_pbe = {}
    for _event in _EVENT_ORDER_PBE:
        _est_pbe[_event] = (_cum_settled_by_event_pbe[_event] * _ratio_pbe) if _ratio_pbe is not None else None
    _month_label = _dt(_mk[0], _mk[1], 1).strftime("%B %Y")
    pbe_rows.append((_month_label, dict(_cum_settled_by_event_pbe), _cum_settled_total_pbe, _cum_batch_pbe,
                      _ratio_pbe, _est_pbe))

ws = wb.create_sheet("Payout by Event")
_pbe_summary = (
    "Estimates how much of the ACTUAL bank-side payout (see Payouts tab, sourced from Payments Hub) is "
    "attributable to each event, as a CUMULATIVE running total by calendar month since August 2026 (September's "
    "figures include both September and August, and so on -- each row is cumulative-to-date, not an isolated "
    "single month). SCOPE: based only on Authorize.net-settled card payments, classified by the event named in "
    "each transaction's Invoice Description (same classification used on the Transactions tab). Venmo-paid "
    "signups/orders never flow through Payments Hub, so they have no bank payout record to allocate and are NOT "
    "included here -- see the Signups/Donations/Store Purchases tabs for the full picture including Venmo. Each "
    "event's estimated payout = that event's cumulative settled $ x ((cumulative actual bank batch volume / "
    "cumulative settled $ across all events). This is an approximation, not an exact figure, since Payments "
    "Hub's own daily report has no per-event breakdown. Added Oct 15, 2026, per Michal: the \"Cumulative Actual "
    "Bank Payout\" figure below is net of refunds -- Payments Hub's daily \"Total Batch Volume\" is gross and "
    "reports that day's refunds separately, so each day's refunds are subtracted before this monthly total is "
    "built, to match the already-refund-netted Authorize.net settled totals it's compared against (same fix as "
    "the Payouts tab)."
)
ws["A1"] = _pbe_summary
ws["A1"].font = SUB_FONT
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A1:L1")
ws.row_dimensions[1].height = 110

headers = ["Month", "Cumulative Settled -- Annual Membership", "Cumulative Settled -- Opening Seminar",
           "Cumulative Settled -- Yom Tzofe", "Cumulative Settled -- Tapuach Store",
           "Cumulative Settled -- Total (Authorize.net)",
           "Cumulative Actual Bank Payout (Payments Hub, net of refunds)", "Allocation Ratio (Payout / Settled)",
           "Est. Payout -- Annual Membership", "Est. Payout -- Opening Seminar", "Est. Payout -- Yom Tzofe",
           "Est. Payout -- Tapuach Store"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

for i, (_month_label, _cum_settled, _cum_settled_total, _cum_batch_val, _ratio, _est) in enumerate(pbe_rows):
    rownum = 3 + i
    vals = [
        _month_label,
        _cum_settled[ANNUAL], _cum_settled[SEMINAR], _cum_settled[YOM_TZOFE], _cum_settled[STORE], _cum_settled_total,
        _cum_batch_val,
        _ratio if _ratio is not None else "--",
        _est[ANNUAL] if _est[ANNUAL] is not None else "--",
        _est[SEMINAR] if _est[SEMINAR] is not None else "--",
        _est[YOM_TZOFE] if _est[YOM_TZOFE] is not None else "--",
        _est[STORE] if _est[STORE] is not None else "--",
    ]
    body_row(ws, rownum, vals, len(vals), start_col=1)
    for col in (2, 3, 4, 5, 6, 7, 9, 10, 11, 12):
        c = ws.cell(row=rownum, column=col)
        if isinstance(c.value, (int, float)):
            c.number_format = CURRENCY
    _rc = ws.cell(row=rownum, column=8)
    if isinstance(_rc.value, (int, float)):
        _rc.number_format = '0.0%'

last_row = 2 + len(pbe_rows)
if not pbe_rows:
    ws.cell(row=3, column=1,
             value="Nothing to show yet -- needs both authorize_report.csv and payouts_*.csv data from August 2026 onward.").font = SUB_FONT
    last_row = 3

ws.freeze_panes = "A3"
ws.auto_filter.ref = f"A2:L{max(last_row, 2)}"

# --- Month-over-month breakdown (added Sep 11, 2026, per Michal) --------------------------------------------
# Same underlying event classification/scope as the cumulative table above, but TRANSPOSED (one column per
# calendar month instead of one row) and showing each month's OWN isolated activity -- NOT a cumulative running
# total -- so it's the natural at-a-glance complement to the cumulative table: read across a row here to see
# month-to-month swings for one event/metric, instead of reading down a column above to see the running total.
mom_col_count = 2 + len(_pbe_sorted_months)  # Metric + Total + one column per month
if pbe_rows:
    mom_title_row = last_row + 2
    mom_header_row = mom_title_row + 1
    ws.cell(row=mom_title_row, column=1,
            value="Month-over-month breakdown -- each month column below is that month's OWN activity in isolation, not cumulative-to-date. Total sums all months shown.").font = SUB_FONT
    ws.merge_cells(start_row=mom_title_row, start_column=1, end_row=mom_title_row, end_column=mom_col_count)

    mom_headers = ["Metric", "Total"] + [_row[0] for _row in pbe_rows]
    for i, h in enumerate(mom_headers):
        ws.cell(row=mom_header_row, column=1 + i, value=h)
    style_header(ws, mom_header_row, len(mom_headers), start_col=1)

    mom_row_defs = [
        ("Annual Membership", lambda _mk: _settled_by_event_month.get((ANNUAL, _mk), 0.0)),
        ("Opening Seminar", lambda _mk: _settled_by_event_month.get((SEMINAR, _mk), 0.0)),
        ("Yom Tzofe", lambda _mk: _settled_by_event_month.get((YOM_TZOFE, _mk), 0.0)),
        ("Tapuach Store", lambda _mk: _settled_by_event_month.get((STORE, _mk), 0.0)),
        ("Total Settled", lambda _mk: sum(_settled_by_event_month.get((_e, _mk), 0.0) for _e in _EVENT_ORDER_PBE)),
        ("Total Actual Bank Payout", lambda _mk: _batch_by_month.get(_mk, 0.0)),
    ]
    for j, (_label, _fn) in enumerate(mom_row_defs):
        rownum = mom_header_row + 1 + j
        _month_vals = [_fn(_mk) for _mk in _pbe_sorted_months]
        vals = [_label, sum(_month_vals)] + _month_vals
        body_row(ws, rownum, vals, len(vals), start_col=1)
        for col in range(2, 1 + mom_col_count):
            c = ws.cell(row=rownum, column=col)
            if isinstance(c.value, (int, float)):
                c.number_format = CURRENCY
        ws.cell(row=rownum, column=2).font = Font(bold=True)
    mom_last_row = mom_header_row + len(mom_row_defs)
else:
    mom_last_row = last_row

autofit(ws, [26, 16] + [20] * max(len(_pbe_sorted_months), 9))

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("payout-by-event months:", len(pbe_rows), "| month-over-month table rows:", (mom_last_row - (last_row + 3)) if pbe_rows else 0)

# =========================================================================
# SHEET: Venmo (Aug 31, 2026, per Michal) -- every Venmo transaction on the org's account since Aug 1, 2026, both
# signup-related AND general org activity, since Venmo is also used day-to-day for unrelated sending/receiving
# (lunch runs, reimbursements, Instacart, gift collections, etc.) -- Michal wanted all of it visible here, not just
# signups. Pulled from Gmail into data/venmo_activity_<daterange>.csv files each refresh -- same accumulating,
# never-edit-an-existing-file staging pattern as the thread_*.csv files for Authorize.net receipts (glob-matched
# and merged below, deduped on exact-row match as a safety net in case a re-run's date range overlaps a prior
# one) -- built outside this script, this just parses whatever's been staged. Treated as optional, like
# authorize_report.csv, since there may be nothing new to pull on a given day.
# Columns: Payee/Payer Name (the other party on the transaction), Comment (the payer's own note if they left one,
# otherwise the merchant/event description Venmo shows), Direction (Incoming/Outgoing), and a signed Amount
# (positive = money in, negative = money out) -- exactly as Michal asked for.
# Signup payments made via Venmo are matched separately (by participant name, from the "Venmo signups for
# tracker" sheet) and already appear on the Transactions tab; this tab is the full raw ledger for reference and
# doesn't feed into Signups/Dashboard math at all, since most of it isn't signup-related.
# =========================================================================
_venmo_activity_files = sorted(glob.glob(os.path.join(DATA_DIR, "venmo_activity*.csv")))
venmo_activity_rows = []
venmo_activity_available = len(_venmo_activity_files) > 0
_seen_venmo_activity_rows = set()
for _va_path in _venmo_activity_files:
    with open(_va_path, newline="") as f:
        _va_rows = list(csv.DictReader(f))
    for _var in _va_rows:
        _va_date = (_var.get("date") or "").strip()
        _va_dir = (_var.get("direction") or "").strip().lower()
        _va_who = (_var.get("counterparty") or "").strip()
        _va_comment = (_var.get("comment") or "").strip()
        try:
            _va_amt = float(_var.get("amount") or 0)
        except ValueError:
            continue
        _dedupe_key = (_va_date, _va_dir, _va_who, _va_comment, _va_amt)
        if _dedupe_key in _seen_venmo_activity_rows:
            continue
        _seen_venmo_activity_rows.add(_dedupe_key)
        _signed = _va_amt if _va_dir == "in" else -_va_amt
        venmo_activity_rows.append([_va_date, _va_who, _va_comment, "Incoming" if _va_dir == "in" else "Outgoing", _signed])
venmo_activity_rows.sort(key=lambda r: r[0])

ws = wb.create_sheet("Venmo")
if venmo_activity_available and venmo_activity_rows:
    _v_summary = (f"Every Venmo transaction on the org's account since Aug 1, 2026 ({len(venmo_activity_rows)} total) -- general "
                  "org activity as well as signups, since Venmo is also used for day-to-day sending/receiving. Signup payments "
                  "made via Venmo are matched by participant name (from Michal's \"Venmo signups for tracker\" sheet) and already "
                  "appear on the Transactions tab; this tab is the full raw ledger for reference and isn't otherwise reconciled.")
elif venmo_activity_available:
    _v_summary = "No Venmo transactions found since Aug 1, 2026."
else:
    _v_summary = ("Venmo activity hasn't been pulled yet (no data/venmo_activity*.csv files found) -- once Gmail is checked for "
                  "Venmo transaction emails, this tab will list them here.")
ws["A1"] = _v_summary
ws["A1"].font = SUB_FONT
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A1:E1")
ws.row_dimensions[1].height = 40

headers = ["Date", "Payee / Payer Name", "Comment", "Direction", "Amount"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

for i, row in enumerate(venmo_activity_rows):
    rownum = 3 + i
    body_row(ws, rownum, row, len(row), start_col=1)
    amt_cell = ws.cell(row=rownum, column=5)
    amt_cell.number_format = CURRENCY
    amt_cell.fill = GREEN_FILL if row[4] >= 0 else RED_FILL

last_row = 2 + len(venmo_activity_rows)
if not venmo_activity_rows:
    ws.cell(row=3, column=1, value="Nothing to show yet -- see note above.").font = SUB_FONT
    last_row = 3

ws.freeze_panes = "A3"
autofit(ws, [12, 26, 40, 12, 14])
ws.auto_filter.ref = f"A2:E{max(last_row, 2)}"

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("venmo activity rows written:", len(venmo_activity_rows), "| available:", venmo_activity_available)

# =========================================================================
# SHEET 5: Signups (rollup)
# =========================================================================
# Early Bird sibling-discount tiers for Annual Membership (per Michal, Aug 21, 2026): price is per FAMILY,
# stacked by how many kids are registering, not a flat per-kid price. All transactions logged so far are
# Early Bird (before the Sep 25, 2026 switch to Regular pricing).
SIBLING_TIERS_EARLYBIRD = [740, 710, 680, 650]  # 1st, 2nd, 3rd, 4th kid

def annual_family_total(email):
    """Sum of sibling-discount tiers for however many roster kids match this payer's email."""
    n = len(dedupe(email_to_kids.get(email.lower(), [])))
    if n == 0 or n > len(SIBLING_TIERS_EARLYBIRD):
        return None
    return sum(SIBLING_TIERS_EARLYBIRD[:n])

def sibling_note(email, paid_note_suffix=""):
    n = len(dedupe(email_to_kids.get(email.lower(), [])))
    total = annual_family_total(email)
    tiers = "+".join(f"${t}" for t in SIBLING_TIERS_EARLYBIRD[:n])
    return f"{n} kid(s) on file -- Early Bird sibling pricing {tiers} = ${total}." + paid_note_suffix

def seminar_note(email, paid_note_suffix=""):
    key = email.lower()
    kids = dedupe(email_to_seminar_kids.get(key, []))
    total = seminar_family_fee.get(key)
    if key in seminar_family_donation:
        donation_amt, base_total, raw_total = seminar_family_donation[key]
        return (f"Seminar Roster confirms {', '.join(kids)}. Its listed Participant Fee (${raw_total:.2f}) isn't a "
                f"valid standalone Seminar price (Price Reference only lists $345 or $305 per participant) -- it "
                f"bakes in a ${donation_amt:.0f} voluntary donation on top of the real ${base_total:.2f} fee. "
                f"Expected Total here is ${base_total:.2f} (excludes the donation); the ${donation_amt:.0f} is "
                f"tracked separately as a Donation Add-on on the Transactions tab, logged as soon as there's a "
                f"payment for this signup regardless of whether more installments are still due.") + paid_note_suffix
    return f"Seminar Roster confirms {', '.join(kids)} -- Participant Fee totals ${total:.2f}." + paid_note_suffix

def yomtzofe_note(email, paid_note_suffix=""):
    key = email.lower()
    kids = dedupe(email_to_yomtzofe_kids.get(key, []))
    total = yomtzofe_family_fee.get(key)
    return (f"Yom Tzofe Roster confirms {', '.join(kids)} -- flat ${YOM_TZOFE_FEE:.0f}/participant totals "
            f"${total:.2f}.") + paid_note_suffix

# =========================================================================
# Scholarships -- read from the "Scholarships 2026-2027" Google Sheet (Michal, Aug 28, 2026): one row per
# family + event, with either a % or a $ discount off the normal price (Type column says which). Families are
# recorded by parent name only, not email, so they're matched here the same way as the Sharon Blumenkranz
# work-email case -- via the payer-name lookup built above (name_to_family_key / name_to_seminar_family_key).
# Some scholarship families haven't registered a kid for one or both events yet; those show as "not yet
# registered" on the Scholarships tab rather than guessing at a dollar amount.
SCHOLARSHIP_NAME_HINTS = {
    "uri klien": ("Doesn't match any parent name on the Roster or Seminar Roster exactly. Possibly a spelling "
                  "variant of \"Ori Klein\" -- the Klein family (Aya, 7th Grade; Tamar, 5th Grade) is on the Annual "
                  "Roster with parents Anat Klein / Ori Klein. Worth confirming before assuming this family hasn't "
                  "registered."),
}

# The Opening Seminar is grades 10-12 only, while Annual Membership covers every grade (Michal, Aug 28, 2026) -- so
# a Seminar scholarship for a family whose kids are all younger isn't "not yet registered," it's simply not
# relevant yet, and shouldn't clutter the Scholarships tab with a row implying something's missing. Checked against
# the Annual Roster (which covers all grades, unlike the Seminar Roster) rather than assumed from the Seminar side.
SEMINAR_GRADES = {10, 11, 12}

def _grade_num(kid_label):
    m = re.match(r"(\d+)", kid_label.rsplit("(", 1)[-1].strip())
    return int(m.group(1)) if m else None

def _family_seminar_eligible(name_key):
    """True if this family (matched by parent name on the Annual Roster) has at least one kid in grades 10-12 --
    i.e. could ever register for the Seminar. False if matched but every kid is younger. None if the family isn't
    found on the Annual Roster at all, so eligibility can't be determined."""
    annual_email = name_to_family_key.get(name_key)
    if not annual_email:
        return None
    kids = dedupe(email_to_kids.get(annual_email, []))
    if not kids:
        return None
    grades = [_grade_num(k) for k in kids]
    return any(g in SEMINAR_GRADES for g in grades if g is not None)

with open(os.path.join(DATA_DIR, "scholarships_refresh.csv")) as f:
    scholarship_rows = list(csv.DictReader(f))

def _resolve_scholarship_family(name, event):
    """(lookup_email, kids, pre_discount_total) for a scholarship's parent name, or (None, [], None) if that
    family hasn't registered a kid for this event yet (or the name doesn't match anything on file)."""
    name_key = name.strip().lower()
    if event == ANNUAL:
        email = name_to_family_key.get(name_key)
        if not email:
            return None, [], None
        return email, dedupe(email_to_kids.get(email, [])), annual_family_total(email)
    elif event == YOM_TZOFE:
        # Added Oct 5, 2026, per Michal -- Yom Tzofe scholarships work the same way, just against the Yom Tzofe
        # Roster/family-fee lookups instead of the Annual/Seminar ones.
        email = name_to_yomtzofe_family_key.get(name_key)
        if not email:
            return None, [], None
        return email, dedupe(email_to_yomtzofe_kids.get(email, [])), yomtzofe_family_fee.get(email)
    else:
        email = name_to_seminar_family_key.get(name_key)
        if not email:
            return None, [], None
        return email, dedupe(email_to_seminar_kids.get(email, [])), seminar_family_fee.get(email)

def _family_alias_emails(kids, event):
    """Every email tied to this family (both parents, plus the kid's own email for Annual) -- not just the one
    canonical email _resolve_scholarship_family happened to return. Needed because a payment might come in under
    EITHER parent's email (same class of bug fixed earlier for Danielle Polak/Sharon Blumenkranz), and the
    scholarship needs to apply no matter which one shows up on the actual transaction."""
    aliases = set()
    if event == ANNUAL:
        rows, name_col, email_cols = roster_rows, "Contact Name", ("Parent 1 Email", "Parent 2 Email", "Email")
    elif event == YOM_TZOFE:
        # Added Oct 5, 2026, per Michal -- yomtzofe_refresh.csv only has one parent-email column (see STEP 1b),
        # plus the registrant's own Email, same two aliases the "registered but not yet paid" backfill above
        # already checks.
        rows, name_col, email_cols = yomtzofe_roster_rows, "Participant Name", ("Parent 1 Email", "Email")
    else:
        rows, name_col, email_cols = seminar_roster_rows, "Participant Name", ("Parent 1 Email", "Parent 2 Email")
    for r in rows:
        label = r[name_col] + " (" + r["Grade This Year"] + ")"
        if label in kids:
            for col in email_cols:
                e = r[col].strip().lower().replace("\\_", "_")
                if e:
                    aliases.add(e)
    return aliases

scholarships = []           # rows for the new Scholarships tab
scholarship_expected = {}  # (lookup_email, signup) -> (discounted_total, note) -- overrides normal pricing
for srow in scholarship_rows:
    name = srow["Family Name"].strip()
    # Added Oct 5, 2026, per Michal -- she added a third ("Yom Tzofe") section to the Scholarships 2026-2027
    # sheet, so this can no longer be a plain binary Annual/Seminar choice (everything that wasn't "Annual..."
    # used to fall through to SEMINAR by default, which would have silently mis-tagged every Yom Tzofe row as a
    # Seminar scholarship instead).
    _evt_raw = srow["Event"].strip().lower()
    if _evt_raw.startswith("annual"):
        event = ANNUAL
    elif "yom tzofe" in _evt_raw or "yomtzofe" in _evt_raw:
        event = YOM_TZOFE
    else:
        event = SEMINAR
    disc_type = srow["Type"].strip().lower()
    disc_raw = srow["Discount"].strip()

    email, kids, pre_total = _resolve_scholarship_family(name, event)

    if email is None:
        if event == SEMINAR and _family_seminar_eligible(name.strip().lower()) is False:
            # Matched on the Annual Roster, but no kid is grades 10-12 -- not a candidate for the Seminar yet, so
            # skip this row entirely rather than implying a missing registration.
            continue
        # Oct 5, 2026: name the specific roster this event actually checks, now that there are three of them.
        _roster_name = {ANNUAL: "Roster", SEMINAR: "Seminar Roster", YOM_TZOFE: "Yom Tzofe Roster"}[event]
        hint = SCHOLARSHIP_NAME_HINTS.get(name.lower(), f"No matching parent name found on the {_roster_name} for this event -- likely just hasn't registered a kid yet.")
        disc_label_unmatched = disc_raw if disc_raw else "(not entered)"
        scholarships.append({"name": name, "event": event, "disc_label": disc_label_unmatched, "kids": "", "n_kids": None,
                              "pre_total": None, "scholarship_amt": None, "owes": None,
                              "status": "Not yet registered", "detail": hint})
        continue

    if not disc_raw:
        # Discount % or $ hasn't been filled in yet on the Scholarships 2026-2027 sheet for this family/event --
        # added Sep 22, 2026 ( ported from the Drive backup's Sep-22 revision on Oct 1, 2026), after a blank
        # Discount cell (Type=Percentage, no value entered) crashed the build (float("") raised ValueError).
        # Rather than guessing a discount amount, flag it the same way an unresolvable price is already flagged
        # below, so Michal can fill it in.
        scholarships.append({"name": name, "event": event, "disc_label": "(not entered)", "kids": "; ".join(kids),
                              "n_kids": len(kids), "pre_total": pre_total, "scholarship_amt": None, "owes": None,
                              "status": "Registered, discount not yet entered",
                              "detail": "This family has a row on the Scholarships 2026-2027 sheet for this event, but the Discount column is blank -- please enter a % or $ amount."})
        continue

    # Oct 4, 2026 (automated daily run, defensible-default fix for a newly-discovered data anomaly -- see Read Me):
    # the live "Scholarships 2026-2027" sheet's Type dropdown now reads "Perentage" (typo for "Percentage") for
    # every row, so the original `disc_type.startswith("perc")` check (which matched "Percentage"/"percent") no
    # longer matches anything and every percentage-style row was falling into the flat-dollar-amount branch below,
    # crashing on float("100.00%"). Rather than guess at or hand-correct the sheet's typo, detect a percentage the
    # same way a human reading the Discount column would: an actual "%" character in the Discount value itself is
    # unambiguous regardless of what the Type column says, so that's added as an OR alongside the original
    # startswith("perc") check (kept so a future correctly-spelled "Percentage" Type with a plain-number Discount,
    # no "%" sign, still works exactly as before).
    pct = float(disc_raw.replace("%", "")) if ("%" in disc_raw or disc_type.startswith("perc")) else None
    disc_label = f"{pct:.0f}%" if pct is not None else (
disc_raw if disc_raw.startswith("$") else f"${disc_raw}")

    if pre_total is None:
        scholarships.append({"name": name, "event": event, "disc_label": disc_label, "kids": "; ".join(kids),
                              "n_kids": len(kids), "pre_total": None, "scholarship_amt": None, "owes": None,
                              "status": "Registered, price unclear",
                              "detail": "Couldn't compute the normal price for this family/event (more kids than the sibling-tier table covers, or similar) -- see the Signups tab note for this family."})
        continue

    if pct is not None:
        owes = round(pre_total * (1 - pct / 100), 2)
    else:
        flat = float(disc_raw.replace("$", "").replace(",", ""))
        owes = max(0.0, round(pre_total - flat, 2))
    scholarship_amt = round(pre_total - owes, 2)
    scholarships.append({"name": name, "event": event, "disc_label": disc_label, "kids": "; ".join(kids),
                          "n_kids": len(kids), "pre_total": pre_total, "scholarship_amt": scholarship_amt,
                          "owes": owes, "status": "Registered", "detail": ""})
    note = (f"Scholarship on file ({disc_label} off, via the Scholarships 2026-2027 sheet): normal price "
            f"${pre_total:.2f} for {len(kids)} kid(s) ({'; '.join(kids)}), reduced by ${scholarship_amt:.2f} to "
            f"${owes:.2f}.")
    for alias in _family_alias_emails(kids, event) | {email}:
        scholarship_expected[(alias, event)] = (owes, note)

print("scholarship rows loaded:", len(scholarship_rows), "| registered families matched:",
      sum(1 for s in scholarships if s["status"] == "Registered"))

# =========================================================================
# Withdrawals, part 2 -- added Oct 6, 2026, per Michal (see "Withdrawals, part 1" above for the helper sheet's
# columns and the roster-removal convention). Now that tx_rows/resolved_email exist, this builds the full detail:
# resolving each withdrawal row's parent name to a family email (same name-matching convention as Scholarships
# above), cross-checking any logged refund against the ACTUAL Authorize.net credit receipt(s) found on the
# Transactions tab (per Michal: don't trust a flat "$30 fee" assumption -- go by the real credited amount, which
# can be split across more than one credit since Authorize.net won't let a single credit exceed the original
# charge), and building the Signups tab Status/Notes override (withdrawal_overrides) and the new Withdrawals tab
# rows (withdrawals_tab_rows). "Updated installment amount" is informational only (Michal, Oct 6, 2026): it's the
# new PER-installment dollar figure she sets directly in Authorize.net for a remaining sibling's plan (the
# installment COUNT doesn't change) -- it doesn't override Expected Total or any other formula here.
# =========================================================================
def _participant_name_matches(participant_name, kid_label):
    """True if every word in `participant_name` (e.g. "Carmel Packin") appears somewhere in `kid_label` (e.g.
    "Packin, Carmel (2nd Grade)") -- word-order-independent, since the helper sheet is typed First Last while
    the roster stores Last, First."""
    kid_lower = kid_label.lower()
    words = [w for w in re.split(r"[\s,]+", participant_name.lower()) if w]
    return bool(words) and all(w in kid_lower for w in words)

# Fallback email resolution for a withdrawal row's parent name when the CURRENT roster has nothing for them --
# added Oct 2026 after testing against real withdrawal data found this was actually the common case, not the
# exception: once every one of a family's kids has withdrawn, Michal removes them from the live roster entirely
# (the documented convention), at which point name_to_family_key/name_to_seminar_family_key/
# name_to_yomtzofe_family_key -- built only from CURRENT roster rows -- have nothing left for that parent name at
# all, so _resolve_withdrawal_family below would always return (None, []) for exactly the families most likely to
# need it. Their Signups combo row still exists (combos are built from Transaction history, not the roster), so
# this falls back to the same email that family's own past Transactions already resolved to (resolved_email,
# keyed by txn_id, with the raw payer email as its own fallback -- see the Transactions-matching loop above).
_name_to_resolved_email_by_signup = defaultdict(dict)  # signup -> {payer name, lowercased: resolved email}
for _txrow in tx_rows:
    _payer_name_key = (_txrow[4] or "").strip().lower()
    if _payer_name_key:
        _name_to_resolved_email_by_signup[_txrow[2]].setdefault(
            _payer_name_key, resolved_email.get(_txrow[8], _txrow[5].lower()))

def _resolve_withdrawal_family(parent_name, signup):
    """(email, kids) for a withdrawal row's parent name -- first tries the relevant CURRENT roster (same
    name_to_*_family_key lookup Scholarships uses above; `kids` here are the ones still on file, used only to
    flag a participant name that no longer matches). If that roster has nothing for this parent (the normal
    state once every one of their kids has already withdrawn and been removed), falls back to the email that
    parent's own past Transactions resolved to, so the Signups override/credit cross-check below still has
    something to match against -- `kids` is correctly [] in that case, since there's genuinely nothing left on
    the roster to list. Returns (None, []) if neither finds anything."""
    name_key = parent_name.strip().lower()
    if signup == ANNUAL:
        email = name_to_family_key.get(name_key)
        kids = dedupe(email_to_kids.get(email, [])) if email else []
    elif signup == YOM_TZOFE:
        email = name_to_yomtzofe_family_key.get(name_key)
        kids = dedupe(email_to_yomtzofe_kids.get(email, [])) if email else []
    else:
        email = name_to_seminar_family_key.get(name_key)
        kids = dedupe(email_to_seminar_kids.get(email, [])) if email else []
    if not email:
        email = _name_to_resolved_email_by_signup.get(signup, {}).get(name_key)
        kids = []
    return email, kids

def _find_matching_withdrawal_credits(email, signup, w_date):
    """Credit (negative-amount) rows on the Transactions tab for this family/signup within a window around the
    withdrawal date -- 3 days before (in case the helper-sheet date lags slightly behind when the credit actually
    posted) to 30 days after (a big refund is sometimes split into several credits over time, since a single
    credit can't exceed the original charge)."""
    from datetime import timedelta
    lo, hi = w_date - timedelta(days=3), w_date + timedelta(days=30)
    out = []
    for row in tx_rows:
        if row[2] != signup or row[3] >= 0:
            continue  # wrong signup, or not a credit (credits are logged as negative amounts -- see Transactions build)
        if resolved_email.get(row[8], row[5]) != email:
            continue
        try:
            d = datetime.strptime(row[0], "%Y-%m-%d")
        except ValueError:
            continue
        if lo <= d <= hi:
            out.append(row)
    return out

_ROSTER_LABEL = {ANNUAL: "Roster", SEMINAR: "Seminar Roster", YOM_TZOFE: "Yom Tzofe Roster"}

# Pass 1 -- parse every raw row once and resolve its family, so rows can be grouped by (email, signup) before
# any credit-matching happens. The grouping matters because a family can log MORE THAN ONE withdrawing
# participant against the same event sharing the SAME original combined charge -- e.g. a parent on an
# installment plan withdrawing 2 kids, where Authorize.net then issues the refund as its own pool of credits
# that doesn't split evenly 1-per-child (added Oct 2026, per Michal: "each have their 30$ deduction from the
# refund... but there is one receipt for both kids"). Matching credits independently per row (the old behavior)
# finds the SAME family-wide credits for every row in the group and compares each one to its own individual
# logged amount -- producing a false "mismatch" (and double-counting the credit total in the Withdrawals tab
# footer) even when the family's combined logged refunds exactly match the family's combined actual credits.
# Grouping first lets the comparison happen at the level the money actually moves at: the family+event, not
# the individual child.
_parsed_withdrawals = []
for _wr in withdrawal_rows_raw:
    parent_name = (_wr.get("parent name") or "").strip()
    participant_name = (_wr.get("participant name") or "").strip()
    rtype_raw = (_wr.get("refund type") or "").strip()
    if not parent_name or not rtype_raw:
        continue  # blank/incomplete row -- skip rather than guess (same as part 1 above)
    rtype_lower = rtype_raw.lower()
    signup = _withdrawal_event_to_signup(_wr.get("event"))
    w_date_raw = (_wr.get("Date") or "").strip()
    try:
        w_date = datetime.strptime(w_date_raw, "%Y-%m-%d")
    except ValueError:
        try:
            w_date = datetime.strptime(w_date_raw, "%m/%d/%Y")
        except ValueError:
            w_date = BUILD_DATE  # unparseable/blank date -- fall back to today rather than dropping the row
    updated_install_raw = (_wr.get("updated installment amount") or "").strip()
    refund_amt_raw = (_wr.get("refund amount") or "").strip()
    is_refund = rtype_lower in ("refund", "both")
    is_adjust = rtype_lower in ("adjust", "both")
    email, family_kids = _resolve_withdrawal_family(parent_name, signup)
    refund_amt_expected = None
    if refund_amt_raw:
        try:
            refund_amt_expected = float(refund_amt_raw.replace("$", "").replace(",", ""))
        except ValueError:
            pass  # leave blank rather than guess at an unparseable figure
    _parsed_withdrawals.append({
        "parent_name": parent_name, "participant_name": participant_name, "rtype_raw": rtype_raw,
        "signup": signup, "w_date": w_date, "updated_install_raw": updated_install_raw,
        "refund_amt_expected": refund_amt_expected, "is_refund": is_refund, "is_adjust": is_adjust,
        "email": email, "family_kids": family_kids,
    })

# Pass 2 -- group refund-eligible, family-matched rows by (email, signup) and find each group's shared
# Authorize.net credits ONCE, using a window wide enough to cover every row in the group (3 days before the
# earliest withdrawal date in the group to 30 days after the latest -- same margins _find_matching_withdrawal_
# credits uses for a single row).
from collections import defaultdict
from datetime import timedelta
_refund_group_dates = defaultdict(list)
for _p in _parsed_withdrawals:
    if _p["is_refund"] and _p["email"]:
        _refund_group_dates[(_p["email"], _p["signup"])].append(_p["w_date"])

_refund_group_info = {}  # (email, signup) -> {"credit_rows": [...], "credit_total": float}
for _gkey, _dates in _refund_group_dates.items():
    _g_email, _g_signup = _gkey
    _lo, _hi = min(_dates) - timedelta(days=3), max(_dates) + timedelta(days=30)
    _rows = []
    for _row in tx_rows:
        if _row[2] != _g_signup or _row[3] >= 0:
            continue
        if resolved_email.get(_row[8], _row[5]) != _g_email:
            continue
        try:
            _d = datetime.strptime(_row[0], "%Y-%m-%d")
        except ValueError:
            continue
        if _lo <= _d <= _hi:
            _rows.append(_row)
    _refund_group_info[_gkey] = {"credit_rows": _rows, "credit_total": sum(-_r[3] for _r in _rows)}

_refund_group_size = {_gkey: len(_dates) for _gkey, _dates in _refund_group_dates.items()}
_refund_group_expected_total = {}
_refund_group_any_missing_expected = {}
for _gkey in _refund_group_dates:
    _g_email, _g_signup = _gkey
    _members = [p for p in _parsed_withdrawals if p["is_refund"] and p["email"] == _g_email and p["signup"] == _g_signup]
    _refund_group_any_missing_expected[_gkey] = any(p["refund_amt_expected"] is None for p in _members)
    _refund_group_expected_total[_gkey] = sum(p["refund_amt_expected"] for p in _members if p["refund_amt_expected"] is not None)

withdrawal_overrides = {}   # (email, signup) -> {"is_refund": bool, "is_adjust": bool, "notes": [str, ...]}
withdrawals_tab_rows = []   # rendered rows for the new Withdrawals tab
for _p in _parsed_withdrawals:
    parent_name, participant_name = _p["parent_name"], _p["participant_name"]
    rtype_raw, signup, w_date = _p["rtype_raw"], _p["signup"], _p["w_date"]
    updated_install_raw, refund_amt_expected = _p["updated_install_raw"], _p["refund_amt_expected"]
    is_refund, is_adjust, email, family_kids = _p["is_refund"], _p["is_adjust"], _p["email"], _p["family_kids"]

    row_notes = []
    if email is None:
        row_notes.append(f"No matching parent name found on the {_ROSTER_LABEL[signup]} -- can't tie this to a "
                          f"Signups row; check the spelling.")
    elif participant_name and family_kids and not any(_participant_name_matches(participant_name, k) for k in family_kids):
        row_notes.append(f"\"{participant_name}\" isn't listed among this family's kids still on file "
                          f"({'; '.join(family_kids)}) -- likely already removed from the roster.")

    credit_rows_found = []
    credit_amt_found = None
    gkey = (email, signup)
    group_size = _refund_group_size.get(gkey, 0)
    if is_refund and email and group_size > 1:
        # Shared family+event refund pool across multiple withdrawing participants -- compare at the family
        # level, not this row alone, so a legitimately uneven credit split doesn't look like a false mismatch.
        info = _refund_group_info[gkey]
        credit_rows_found = info["credit_rows"]
        group_total_credit = info["credit_total"]
        group_total_expected = _refund_group_expected_total[gkey]
        group_missing = _refund_group_any_missing_expected[gkey]
        others = ", ".join(f"${p['refund_amt_expected']:,.2f}" for p in _parsed_withdrawals
                            if p is not _p and p["is_refund"] and p["email"] == email and p["signup"] == signup
                            and p["refund_amt_expected"] is not None)
        if not credit_rows_found:
            credit_amt_found = 0.0
            row_notes.append(f"No matching Authorize.net credit found yet for this family/event -- refund may "
                              f"still be processing. This family has {group_size} withdrawing participants "
                              f"logged for this event" + (f" (this one: ${refund_amt_expected:,.2f} logged)"
                                                           if refund_amt_expected is not None else "") + ".")
        elif group_missing:
            credit_amt_found = None
            row_notes.append(f"This family has {group_size} withdrawing participants logged for this event, but "
                              f"at least one has no refund amount entered on the helper sheet -- can't verify the "
                              f"family's total against Authorize.net. The family's actual credit receipt(s) for "
                              f"this event total ${group_total_credit:,.2f} ({len(credit_rows_found)} "
                              f"credit{'s' if len(credit_rows_found) != 1 else ''}, "
                              f"{', '.join(r[0] for r in credit_rows_found)}).")
        elif abs(group_total_credit - group_total_expected) <= 0.01:
            credit_amt_found = refund_amt_expected  # this participant's own logged share, confirmed in aggregate
            row_notes.append(f"Refund confirmed as part of a combined family refund: {group_size} withdrawing "
                              f"participants logged ${group_total_expected:,.2f} total for this event (this one: "
                              f"${refund_amt_expected:,.2f}" + (f"; others: {others}" if others else "") + "), "
                              f"matching the ${group_total_credit:,.2f} Authorize.net actually credited across "
                              f"{len(credit_rows_found)} receipt{'s' if len(credit_rows_found) != 1 else ''} "
                              f"({', '.join(r[0] for r in credit_rows_found)}) -- credited as separate, "
                              f"not-necessarily-even amounts rather than one credit per child, per Michal.")
        else:
            credit_amt_found = None
            row_notes.append(f"${group_total_expected:,.2f} logged in total for this family's {group_size} "
                              f"withdrawing participants on this event (this one: ${refund_amt_expected:,.2f}"
                              + (f"; others: {others}" if others else "") + "), but Authorize.net's actual credit "
                              f"receipt(s) for this family/event total ${group_total_credit:,.2f} "
                              f"({len(credit_rows_found)} credit{'s' if len(credit_rows_found) != 1 else ''}, "
                              f"{', '.join(r[0] for r in credit_rows_found)}) -- going by the receipt total, not "
                              f"the logged figures (per Michal, Oct 6, 2026: don't assume a flat fee).")
    elif is_refund and email:
        credit_rows_found = _find_matching_withdrawal_credits(email, signup, w_date)
        credit_amt_found = sum(-r[3] for r in credit_rows_found)
        if not credit_rows_found:
            row_notes.append("No matching Authorize.net credit found yet on the Transactions tab -- refund may "
                              "still be processing." + (f" (${refund_amt_expected:,.2f} logged as expected.)"
                                                         if refund_amt_expected is not None else ""))
        elif refund_amt_expected is not None and abs(credit_amt_found - refund_amt_expected) > 0.01:
            row_notes.append(f"${refund_amt_expected:,.2f} logged as the expected refund, but Authorize.net's "
                              f"actual credit receipt(s) total ${credit_amt_found:,.2f} ({len(credit_rows_found)} "
                              f"credit{'s' if len(credit_rows_found) != 1 else ''}, "
                              f"{', '.join(r[0] for r in credit_rows_found)}) -- going by the receipt total, not "
                              f"the logged figure (per Michal, Oct 6, 2026: don't assume a flat fee).")
        else:
            row_notes.append(f"Refund confirmed: ${credit_amt_found:,.2f} credited ({len(credit_rows_found)} "
                              f"receipt{'s' if len(credit_rows_found) != 1 else ''}, "
                              f"{', '.join(r[0] for r in credit_rows_found)}).")
    elif is_refund:
        row_notes.append("Can't check for a matching Authorize.net credit -- family/event not matched above.")

    if is_adjust:
        if updated_install_raw:
            row_notes.append(f"Future installments updated to ${updated_install_raw.replace('$', '').strip()} "
                              f"each (same installment count as before) -- set directly in Authorize.net, not a "
                              f"workbook formula.")
        else:
            row_notes.append("Marked as an installment adjustment, but no updated amount was entered on the "
                              "helper sheet.")

    full_note = (f"{participant_name or 'Participant'} withdrew {w_date:%b %d, %Y}"
                 + (f" ({parent_name}'s family)" if parent_name else "") + ". " + " ".join(row_notes)).strip()

    withdrawals_tab_rows.append({
        "date": w_date, "parent_name": parent_name, "participant_name": participant_name, "signup": signup,
        "refund_type_raw": rtype_raw, "updated_install_raw": updated_install_raw,
        "refund_amt_expected": refund_amt_expected, "credit_amt_found": credit_amt_found,
        "matched_email": email or "", "note": full_note,
    })

    if email:
        key = (email, signup)
        ov = withdrawal_overrides.setdefault(key, {"is_refund": False, "is_adjust": False, "notes": []})
        ov["is_refund"] = ov["is_refund"] or is_refund
        ov["is_adjust"] = ov["is_adjust"] or is_adjust
        ov["notes"].append(full_note)

print("withdrawal rows loaded:", len(withdrawal_rows_raw), "| matched to a family:",
      sum(1 for r in withdrawals_tab_rows if r["matched_email"]))

# Hand-written overrides for genuine special cases -- everything else defaults to the generic sibling-tier /
# Seminar Roster pricing computed automatically below, so this only needs to grow when a payment is a real oddity
# (test charge, short payment, roster/fee mismatch, duplicate charge, unrelated purchase), not for every payer.
SPECIAL_NOTES = {
    ("kroitero94@gmail.com", ANNUAL): (0, "Verification charge only ($1.00) -- not a real due."),
    ("noa.reikhav@gmail.com", ANNUAL): (annual_family_total("noa.reikhav@gmail.com"),
        sibling_note("noa.reikhav@gmail.com", " Includes an apparent duplicate $207.15 charge on Aug 17 that was refunded (credited) on Aug 19 -- see the per-row notes on Transactions. The two cancel out, leaving one real $207.15 payment.")),
}

# Legacy Annual sibling pricing (Michal, Aug 28, 2026): Tsafrir Oranski, Dikla Yerushalmi, and Michal Sheves all
# signed their kids up while the sibling table was still $730 (1st kid) / $700 (2nd kid), before it was corrected
# to the current $740/$710/$680/$650 table. Their Expected Total should use the price they actually agreed to at
# signup, not today's table -- this is what made their payments look like odd partial amounts (Tsafrir's $730 is
# actually a full, exact payment at the old price; Dikla's $715 is exactly half of the old 2-kid total $1,430).
LEGACY_SIBLING_TIERS_2025 = [730, 700]  # 1st, 2nd kid -- no 3rd/4th-kid legacy price given, but none of these 3
                                         # families have more than 2 kids on file, so it isn't needed.

def _legacy_annual_total(email):
    n = len(dedupe(email_to_kids.get(email.lower(), [])))
    if n == 0 or n > len(LEGACY_SIBLING_TIERS_2025):
        return None
    return sum(LEGACY_SIBLING_TIERS_2025[:n])

for _legacy_email in ("tsafrir@gmail.com", "diklaisaac@yahoo.com", "michaleet@gmail.com"):
    _legacy_total = _legacy_annual_total(_legacy_email)
    if _legacy_total is None:
        continue  # more kids than the 2-tier legacy table covers -- shouldn't happen for these 3, but don't guess
    _n = len(dedupe(email_to_kids.get(_legacy_email, [])))
    _tiers_txt = "+".join(f"${t}" for t in LEGACY_SIBLING_TIERS_2025[:_n])
    SPECIAL_NOTES[(_legacy_email, ANNUAL)] = (_legacy_total,
        f"{_n} kid(s) on file -- signed up under the pre-correction Annual sibling pricing ({_tiers_txt} = "
        f"${_legacy_total}), not the current $740/$710/$680/$650 table (per Michal, Aug 28, 2026: this family "
        f"registered before the price was corrected). Expected Total here uses the price they actually agreed to.")

# Everything else: default to the generic sibling-discount total (Annual) or Seminar Roster Participant Fee
# (Seminar) for whichever payer/signup combos actually show up in the transactions, in the order first seen.
seen = set()
combos = []
expected_map = {}
combo_lookup_email = {}  # (raw payer email, signup) -> roster email actually used to match (see note below)
for row in tx_rows:
    signup = row[2]
    if signup == STORE:
        continue  # not a signup fee -- shown on Transactions, not tracked on Signups
    email = row[5]
    # The roster email actually used to match this payment to a family (may differ from the payer's raw email
    # via the name fallback above, e.g. Sharon Blumenkranz paying from a work email) -- used for family-total /
    # note lookups below, while `email` (the raw payer email) stays what's shown on the Signups tab and what
    # the live COUNTIFS/SUMIFS formulas there match against Transactions, so those formulas keep working.
    lookup_email = resolved_email.get(row[8], email)
    key = (email, signup)
    combo_lookup_email[key] = lookup_email
    if key not in seen:
        seen.add(key)
        if lookup_email != email:
            seen.add((lookup_email, signup))  # so the "registered but not yet paid" check below recognizes this family as paid
        combos.append({"name": row[4], "email": email, "signup": signup, "matched": row[9]})
    if key in expected_map:
        continue
    if (lookup_email, signup) in scholarship_expected:
        expected_map[key] = scholarship_expected[(lookup_email, signup)]
    elif key in SPECIAL_NOTES:
        expected_map[key] = SPECIAL_NOTES[key]
    elif signup == ANNUAL:
        exp = annual_family_total(lookup_email)
        expected_map[key] = ((exp, sibling_note(lookup_email)) if exp is not None
                              else (None, "No clean match to the sibling-discount pricing table (more kids than the table covers, or none on the Roster tab) -- please confirm the agreed total."))
    elif signup == SEMINAR:
        exp = seminar_family_fee.get(lookup_email)
        expected_map[key] = ((exp, seminar_note(lookup_email)) if exp else
                              (None, "No matching entry on the Seminar Roster tab for this email -- please confirm the agreed total."))
    else:  # YOM_TZOFE
        exp = yomtzofe_family_fee.get(lookup_email)
        expected_map[key] = ((exp, yomtzofe_note(lookup_email)) if exp else
                              (None, "No matching entry on the Yom Tzofe Roster tab for this email -- please confirm the agreed total."))

# BUILD_DATE is set once at the top of the script (see above) -- reused here for the "how long has this been
# unpaid" check below.

# Earliest registration timestamp per family, for the "how long has this been unpaid" note below.
family_earliest_reg = {}
for srow in seminar_roster_rows:
    key = (srow["Parent 1 Email"].strip().lower().replace("\\_", "_")
           or srow["Parent 2 Email"].strip().lower().replace("\\_", "_"))
    if not key:
        continue
    try:
        reg_dt = datetime.strptime(srow["Registration Date"], "%Y-%m-%d %H:%M")
    except ValueError:
        continue
    if key not in family_earliest_reg or reg_dt < family_earliest_reg[key]:
        family_earliest_reg[key] = reg_dt

n_no_payment = 0
seen_seminar_family = set()
for srow in seminar_roster_rows:
    p1 = srow["Parent 1 Email"].strip().lower().replace("\\_", "_")
    p2 = srow["Parent 2 Email"].strip().lower().replace("\\_", "_")
    key = p1 or p2
    if not key or key in seen_seminar_family:
        continue
    seen_seminar_family.add(key)
    # A payment may have come in under EITHER parent's email, not just the canonical key above --
    # check both before treating this family as unpaid (see the seminar_family_fee note above).
    if (key, SEMINAR) in seen or (p1 and (p1, SEMINAR) in seen) or (p2 and (p2, SEMINAR) in seen):
        continue  # already has at least one payment logged
    seen.add((key, SEMINAR))
    kids = dedupe(email_to_seminar_kids.get(key, []))
    payer_name = srow["Parent 1 Contact Name"] or srow["Parent 2 Contact Name"]
    combos.append({"name": payer_name, "email": key, "signup": SEMINAR, "matched": "; ".join(kids)})
    reg_dt = family_earliest_reg.get(key)
    hours_ago = (BUILD_DATE - reg_dt).total_seconds() / 3600 if reg_dt else None
    if hours_ago is not None and hours_ago <= 24:
        recency = (f"Registered {reg_dt:%b %d, %I:%M %p} -- only {hours_ago:.0f} hours ago, so this may just be the normal "
                   f"overnight gap before the first installment processes, not a real non-payment.")
    elif reg_dt:
        recency = f"Registered {reg_dt:%b %d, %I:%M %p} -- {hours_ago/24:.1f} days ago with nothing paid, so probably worth a nudge."
    else:
        recency = "Registration date unavailable."
    if (key, SEMINAR) in scholarship_expected:
        exp, sch_note = scholarship_expected[(key, SEMINAR)]
        note = f"{sch_note} No payment logged yet. {recency}"
    else:
        exp = seminar_family_fee.get(key)
        note = f"Registered on the Seminar Roster ({'; '.join(kids)}, ${exp:.2f}) but no payment logged yet. {recency}"
    expected_map[(key, SEMINAR)] = (exp, note)
    n_no_payment += 1
print("seminar signups added with no payment yet:", n_no_payment)

# Same "registered but not yet paid" backfill, for Yom Tzofe -- added Oct 5, 2026. Simpler than Seminar's version
# since there's only one parent-email column to check (Parent 1 Email, falling back to the registrant's own
# Email) and no grade restriction. Scholarships ARE checked below (added later the same day, per Michal), same
# pattern as the Seminar loop above.
_yomtzofe_family_earliest_reg = {}
for yrow in yomtzofe_roster_rows:
    key = (yrow["Parent 1 Email"].strip().lower().replace("\\_", "_")
           or yrow["Email"].strip().lower().replace("\\_", "_"))
    if not key:
        continue
    try:
        reg_dt = datetime.strptime(yrow["Registration Date"], "%Y-%m-%d %H:%M")
    except ValueError:
        continue
    if key not in _yomtzofe_family_earliest_reg or reg_dt < _yomtzofe_family_earliest_reg[key]:
        _yomtzofe_family_earliest_reg[key] = reg_dt

n_no_payment_yomtzofe = 0
seen_yomtzofe_family = set()
for yrow in yomtzofe_roster_rows:
    p1 = yrow["Parent 1 Email"].strip().lower().replace("\\_", "_")
    pe = yrow["Email"].strip().lower().replace("\\_", "_")
    key = p1 or pe
    if not key or key in seen_yomtzofe_family:
        continue
    seen_yomtzofe_family.add(key)
    # A payment may have come in under either the parent's email or the registrant's own email -- check both
    # before treating this family as unpaid (same idea as the Seminar loop above).
    if (key, YOM_TZOFE) in seen or (p1 and (p1, YOM_TZOFE) in seen) or (pe and (pe, YOM_TZOFE) in seen):
        continue  # already has at least one payment logged
    seen.add((key, YOM_TZOFE))
    kids = dedupe(email_to_yomtzofe_kids.get(key, []))
    payer_name = yrow["Parent 1 Contact Name"] or yrow["Participant Name"]
    combos.append({"name": payer_name, "email": key, "signup": YOM_TZOFE, "matched": "; ".join(kids)})
    reg_dt = _yomtzofe_family_earliest_reg.get(key)
    hours_ago = (BUILD_DATE - reg_dt).total_seconds() / 3600 if reg_dt else None
    if hours_ago is not None and hours_ago <= 24:
        recency = (f"Registered {reg_dt:%b %d, %I:%M %p} -- only {hours_ago:.0f} hours ago, so this may just be the normal "
                   f"overnight gap before payment processes, not a real non-payment.")
    elif reg_dt:
        recency = f"Registered {reg_dt:%b %d, %I:%M %p} -- {hours_ago/24:.1f} days ago with nothing paid, so probably worth a nudge."
    else:
        recency = "Registration date unavailable."
    if (key, YOM_TZOFE) in scholarship_expected:
        # Added Oct 5, 2026, per Michal -- she added a Yom Tzofe scholarship section to the Scholarships sheet,
        # so a registered-but-unpaid family here might be a scholarship family owing a reduced amount, not the
        # full $95/kid flat fee (same pattern as the Seminar version of this loop above).
        exp, sch_note = scholarship_expected[(key, YOM_TZOFE)]
        note = f"{sch_note} No payment logged yet. {recency}"
    else:
        exp = yomtzofe_family_fee.get(key)
        note = f"Registered on the Yom Tzofe Roster ({'; '.join(kids)}, ${exp:.2f}) but no payment logged yet. {recency}"
    expected_map[(key, YOM_TZOFE)] = (exp, note)
    n_no_payment_yomtzofe += 1
print("yom tzofe signups added with no payment yet:", n_no_payment_yomtzofe)

# Backfill Annual rows for scholarship families that haven't made any Annual payment yet -- same idea as the
# Seminar "not yet paid" loop above, but scoped to scholarship families only (Bug found Aug 28, 2026: Lilach Vos
# has a 100% Annual scholarship, $0 due, but had never made an Annual payment, so no Annual row existed for her
# at all -- no "$0 due" confirmation anywhere). Not extended to the full Annual roster the way Seminar is, since
# doing that for all ~150 families would bury the sheet in early-season "hasn't paid yet" noise for families
# simply not yet due -- that's still a documented open gap, see Read Me.
n_no_payment_annual_scholarship = 0
for s in scholarships:
    if s["status"] != "Registered" or s["event"] != ANNUAL:
        continue
    kids_list = s["kids"].split("; ") if s["kids"] else []
    aliases = _family_alias_emails(kids_list, ANNUAL)
    if any((a, ANNUAL) in seen for a in aliases):
        continue  # already has at least one Annual payment logged
    key = next(iter(aliases), None)
    if not key:
        continue
    seen.add((key, ANNUAL))
    exp, sch_note = scholarship_expected[(key, ANNUAL)]
    combos.append({"name": s["name"], "email": key, "signup": ANNUAL, "matched": s["kids"]})
    expected_map[(key, ANNUAL)] = (exp, f"{sch_note} No Annual payment logged yet.")
    n_no_payment_annual_scholarship += 1
print("annual scholarship signups added with no payment yet:", n_no_payment_annual_scholarship)

ws = wb.create_sheet("Signups")
ws["A1"] = ("One row per family + signup -- including Seminar families with a confirmed signup but no payment logged yet "
              "Status \"Not yet paid\"). # Payments, Total Paid, and Donations Included all update automatically from the "
            "Transactions tab as new receipts are added there (Total Paid excludes any amount logged as a Donation "
            "Add-on). Fill in the yellow Expected Total cells. Total Installments works the same way as Expected Total: "
            "yellow = confirmed by you or the family; blue = Claude's algorithmic estimate from the payment pattern "
            "(see Read Me) -- worth a sanity check, not guaranteed.")
ws["A1"].font = SUB_FONT
ws.merge_cells("A1:L1")
ws.row_dimensions[1].height = 28

# Total Installments a family agreed to for a signup. Two tiers, shown with different fills:
#  1) confirmed_installments (yellow) -- a family/Michal actually said so. Highest confidence.
#  2) estimated_installments (blue) -- inferred algorithmically (Michal, Aug 28, 2026: figure it out "from the total
#     and what is already paid, removing any donation and calculating in discounts"). For a family still
#     mid-payment, checks whether what they've paid so far (in gross dollars, including any embedded donation)
#     divides evenly into their full agreed total (Expected Total -- already net of any scholarship discount --
#     plus any known embedded donation) for some small number of equal installments. Only kept when the fit is
#     tight (within a couple cents); otherwise left blank rather than guessed at.
confirmed_installments = {
    ("yaelhal@gmail.com", SEMINAR): 3,  # Brutman/Yonatan -- confirmed by Michal, Aug 23, 2026 (~$123.33 x 3 = $370)
}

_gross_paid_by_combo = defaultdict(float)  # (raw payer email, signup) -> total $ actually charged so far
_n_paid_by_combo = defaultdict(int)
for _row in tx_rows:
    _signup = _row[2]
    if _signup == STORE:
        continue
    _ckey = (_row[5], _signup)
    _gross_paid_by_combo[_ckey] += _row[3]
    _n_paid_by_combo[_ckey] += 1

estimated_installments = {}  # (raw payer email, signup) -> n
for _combo in combos:
    _ckey = (_combo["email"], _combo["signup"])
    if _ckey in confirmed_installments:
        continue
    _exp, _ = expected_map.get(_ckey, (None, ""))
    if _exp is None or _exp <= 0:
        continue
    _lookup_email = combo_lookup_email.get(_ckey, _combo["email"])
    _donation_amt = 0.0
    if _combo["signup"] == SEMINAR and _lookup_email in seminar_family_donation:
        _donation_amt = seminar_family_donation[_lookup_email][0]
    _gross_expected = _exp + _donation_amt
    _gross_paid = _gross_paid_by_combo.get(_ckey, 0.0)
    _n_paid = _n_paid_by_combo.get(_ckey, 0)
    if _n_paid == 0 or _gross_paid <= 0 or _gross_paid >= _gross_expected - 0.01:
        continue  # nothing paid yet, or already fully/over paid -- nothing to infer
    for _n in range(_n_paid + 1, 13):  # smallest plausible count of EQUAL installments beyond what's paid so far
        _per_installment = _gross_expected / _n
        if abs(_gross_paid - _n_paid * _per_installment) <= max(0.02, 0.01 * _n_paid):
            estimated_installments[_ckey] = _n
            break

print("installment counts: confirmed", len(confirmed_installments), "| estimated", len(estimated_installments))

headers = ["Payer Name", "Payer Email", "Matched Child(ren)", "Signup", "# Payments", "Total Installments\n(if known)",
           "Total Paid\n(excl. donations)", "Donations\nIncluded", "Expected Total\n(enter manually)",
           "Remaining Balance", "Status", "Notes"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

for i, combo in enumerate(combos):
    rownum = 3 + i
    ws.cell(row=rownum, column=1, value=combo["name"])
    ws.cell(row=rownum, column=2, value=combo["email"])
    ws.cell(row=rownum, column=3, value=combo["matched"])
    ws.cell(row=rownum, column=4, value=combo["signup"])
    ws.cell(row=rownum, column=5, value=f"=COUNTIFS(Transactions!$F$3:$F${TX_LAST_ROW},B{rownum},Transactions!$C$3:$C${TX_LAST_ROW},D{rownum})")
    ckey = (combo["email"], combo["signup"])
    if ckey in confirmed_installments:
        install_cell = ws.cell(row=rownum, column=6, value=confirmed_installments[ckey])
        install_cell.fill = YELLOW_FILL
    elif ckey in estimated_installments:
        install_cell = ws.cell(row=rownum, column=6, value=estimated_installments[ckey])
        install_cell.fill = BLUE_FILL
    else:
        ws.cell(row=rownum, column=6, value=None)  # no fill -- nothing known or inferred for this row
    ws.cell(row=rownum, column=7, value=f"=SUMIFS(Transactions!$N$3:$N${TX_LAST_ROW},Transactions!$F$3:$F${TX_LAST_ROW},B{rownum},Transactions!$C$3:$C${TX_LAST_ROW},D{rownum})")
    ws.cell(row=rownum, column=8, value=f"=SUMIFS(Transactions!$M$3:$M${TX_LAST_ROW},Transactions!$F$3:$F${TX_LAST_ROW},B{rownum},Transactions!$C$3:$C${TX_LAST_ROW},D{rownum})")
    exp, note = expected_map.get(ckey, (None, "No clean match to known pricing -- please confirm the agreed total."))
    exp_cell = ws.cell(row=rownum, column=9, value=exp)
    exp_cell.fill = YELLOW_FILL
    ws.cell(row=rownum, column=10, value=f'=IF(I{rownum}="","",I{rownum}-G{rownum})')
    # Withdrawals override (added Oct 6, 2026, per Michal) -- a static Status value instead of the usual live
    # formula, for any family+signup logged on the Withdrawals helper sheet (see "Withdrawals, part 2" above).
    # Checked against BOTH the raw payer email and the resolved family email (combo_lookup_email), the same way
    # the "registered but not yet paid" backfills above treat the two as interchangeable for a given family.
    wov = (withdrawal_overrides.get((combo_lookup_email.get(ckey, combo["email"]), combo["signup"]))
           or withdrawal_overrides.get(ckey))
    if wov:
        status_override = ("Refund + Adjusted (withdrawal)" if (wov["is_refund"] and wov["is_adjust"])
                            else "Refund (withdrawal)" if wov["is_refund"] else "Adjusted (withdrawal)")
        ws.cell(row=rownum, column=11, value=status_override)
    else:
        ws.cell(row=rownum, column=11, value=(f'=IF(I{rownum}="","Enter expected total",'
                                               f'IF(I{rownum}<=0,"Paid in full",'
                                               f'IF(G{rownum}=0,"Not yet paid",'
                                               f'IF(J{rownum}<=0,"Paid in full","In progress (installment plan)"))))'))
    if ckey in estimated_installments:
        note = (f"Estimated {estimated_installments[ckey]} total installments (blue cell) -- what's been paid so far "
                f"divides evenly into the full agreed amount for that count; not confirmed by the family, just a "
                f"pattern match. {note}")
    if wov:
        note = " | ".join(wov["notes"]) + (" " + note if note else "")
    ws.cell(row=rownum, column=12, value=note)
    for c in range(1, 13):
        cell = ws.cell(row=rownum, column=c)
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.alignment = Alignment(vertical="center", wrap_text=(c in (3, 12)))
    ws.cell(row=rownum, column=7).number_format = CURRENCY
    ws.cell(row=rownum, column=8).number_format = CURRENCY
    ws.cell(row=rownum, column=9).number_format = CURRENCY
    ws.cell(row=rownum, column=10).number_format = CURRENCY

last_row = 2 + len(combos)
SIGNUPS_LAST_ROW = last_row  # kept under its own name -- plain `last_row` is reused as scratch by later sheets below,
                                # but the Dashboard tab still needs this exact value to reference the Signups tab correctly.
ws.freeze_panes = "A3"
autofit(ws, [20, 26, 34, 40, 10, 14, 15, 12, 15, 15, 24, 46])
ws.row_dimensions[2].height = 30
ws.auto_filter.ref = f"A2:L{last_row}"

# Conditional formatting on Status column
from openpyxl.formatting.rule import CellIsRule, FormulaRule
ws.conditional_formatting.add(
    f"K3:K{last_row}",
    FormulaRule(formula=[f'$K3="Paid in full"'], fill=GREEN_FILL)
)
ws.conditional_formatting.add(
    f"K3:K{last_row}",
    FormulaRule(formula=[f'$K3="Enter expected total"'], fill=ORANGE_FILL)
)
ws.conditional_formatting.add(
    f"K3:K{last_row}",
    FormulaRule(formula=[f'!$K3="Not yet paid"'], fill=RED_FILL)
)
ws.conditional_formatting.add(
    f"K3:K{last_row}",
    FormulaRule(formula=[f'$K3="In progress (installment plan)"'], fill=YELLOW_FILL)
)
# Withdrawal-driven statuses (added Oct 6, 2026, per Michal) -- see the Withdrawals override above.
ws.conditional_formatting.add(
    f"K3:K{last_row}",
    FormulaRule(formula=[f'OR($K3="Refund (withdrawal)",$K3="Adjusted (withdrawal)",$K3="Refund + Adjusted (withdrawal)")'],
                fill=PURPLE_FILL)
)

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("signups combos:", len(combos))

# =========================================================================
# SHEET: Scholarships
# =========================================================================
ws = wb.create_sheet("Scholarships")
ws["A1"] = ("One row per family + event scholarship, from the \"Scholarships 2026-2027\" sheet (Michal, Aug 28, 2026; a third "
            "section for Yom Tzofe was added Oct 5, 2026). \"Amount They'll Pay\" is already reflected as that family's Expected "
            "Total on the Signups tab. Rows marked \"Not yet registered\" have a scholarship on file but no matching kid on the "
            "Roster, Seminar Roster, or Yom Tzofe Roster yet for that event -- nothing to compute until they sign up. A Seminar "
            "scholarship for a family whose kids are all younger than 10th grade doesn't appear here at all (Seminar is grades "
            "10-12 only -- see Read Me).")
ws["A1"].font = SUB_FONT
ws.merge_cells("A1:I1")
ws.row_dimensions[1].height = 28

headers = ["Family (Parent Name)", "Event", "Discount", "Kids Covered", "# Kids", "Normal Price",
           "Scholarship Amount", "Amount They'll Pay", "Notes"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

EVENT_LABEL = {ANNUAL: "Annual Membership", SEMINAR: "Opening Seminar", YOM_TZOFE: "Yom Tzofe", STORE: "Tapuach Store checkout"}

# Sort each event section so registered/utilized families are grouped together first, with anything carrying a
# Notes comment (not yet registered, price unclear) pushed to the end of that section (Michal, Aug 28, 2026) --
# a stable sort, so the original relative order within each of those two groups is otherwise unchanged.
_event_order = {ANNUAL: 0, SEMINAR: 1, YOM_TZOFE: 2}
scholarships.sort(key=lambda s: (_event_order[s["event"]], s["status"] != "Registered"))

for i, s in enumerate(scholarships):
    rownum = 3 + i
    vals = [s["name"], EVENT_LABEL[s["event"]], s["disc_label"], s["kids"], s["n_kids"],
            s["pre_total"], s["scholarship_amt"], s["owes"], s["detail"]]
    body_row(ws, rownum, vals, len(vals), start_col=1)
    for col in (6, 7, 8):
        ws.cell(row=rownum, column=col).number_format = CURRENCY
    ws.cell(row=rownum, column=9).alignment = Alignment(vertical="center", wrap_text=True)
    if s["status"] != "Registered":
        for c in range(1, 10):
            ws.cell(row=rownum, column=c).fill = ORANGE_FILL

last_data_row = 2 + len(scholarships)
ws.freeze_panes = "A3"
autofit(ws, [22, 18, 10, 34, 8, 14, 16, 16, 50])
ws.auto_filter.ref = f"A2:I{last_data_row}"

trow = last_data_row + 2
ws.cell(row=trow, column=1, value="Totals").font = SECTION_FONT
trow += 1
for event in (ANNUAL, SEMINAR, YOM_TZOFE):
    matched = [s for s in scholarships if s["event"] == event and s["status"] == "Registered"]
    n_fam = len(matched)
    n_kids = sum(s["n_kids"] for s in matched)
    total_sch = sum(s["scholarship_amt"] for s in matched)
    total_owed = sum(s["owes"] for s in matched)
    label = ws.cell(row=trow, column=1,
                     value=f"{EVENT_LABEL[event]} -- {n_fam} scholarship family{'y' if n_fam == 1 else 'ies'}, {n_kids} kid(s)")
    label.font = BOLD_FONT
    ws.cell(row=trow, column=6, value="Scholarship $ granted ->").font = BODY_FONT
    ws.cell(row=trow, column=7, value=total_sch).number_format = CURRENCY
    ws.cell(row=trow, column=7).font = BOLD_FONT
    ws.cell(row=trow, column=8, value=total_owed).number_format = CURRENCY
    ws.cell(row=trow, column=8).font = BOLD_FONT
    trow += 1
n_pending = sum(1 for s in scholarships if s["status"] != "Registered")
if n_pending:
    ws.cell(row=trow, column=1,
            value=f"{n_pending} scholarship row(s) not yet registered / unclear -- see Notes column above.").font = SUB_FONT

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("scholarship tab rows:", len(scholarships), "| not yet registered/unclear:", n_pending)

# =========================================================================
# SHEET: Withdrawals (added Oct 6, 2026, per Michal) -- one row per entry on her "Withdrawals for Payments
# tracker" helper sheet (data/withdrawals_refresh.csv, re-pulled fresh every run like Scholarships/Venmo signups
# -- not backed up to Drive, see STEP 0/STEP 8 in the scheduled task). "Refund Confirmed" is the actual
# Authorize.net credit receipt total found on the Transactions tab, not the logged "Refund Amount" Michal
# expected -- per Michal, Oct 6, 2026, go by the receipt, since a refund can be split across more than one
# credit and isn't reliably a flat amount off. "Updated Installment Amt" is informational only -- the new
# per-installment dollar figure Michal sets directly in Authorize.net for a remaining sibling's plan; it doesn't
# feed into any total here. Rows that couldn't be matched to a parent name on the relevant roster are shaded
# orange. See the Signups tab for the matching Status/Notes override on that family's row, where one exists.
# =========================================================================
ws = wb.create_sheet("Withdrawals")
ws["A1"] = ("One row per entry on the \"Withdrawals for Payments tracker\" helper sheet -- logged by Michal whenever a family "
            "withdraws a participant mid-season. \"Refund Confirmed\" is the actual Authorize.net credit receipt total found on "
            "the Transactions tab, not the logged \"Refund Amount\" Michal expected -- go by the receipt (don't assume a flat "
            "fee off). \"Updated Installment Amt\" is informational only (set directly in Authorize.net); it doesn't change any "
            "total here -- Expected Total on the Signups tab instead relies on the withdrawn participant being removed from the "
            "live Roster/Seminar Roster/Yom Tzofe Roster sheet. The matching Signups tab row's Status/Notes are updated too.")
ws["A1"].font = SUB_FONT
ws.merge_cells("A1:J1")
ws.row_dimensions[1].height = 42

headers = ["Date", "Parent Name", "Participant Name", "Event", "Refund Type",
           "Updated Installment Amt\n(per installment, info only)", "Refund Amount\n(logged/expected)",
           "Refund Confirmed\n(Authorize.net credits)", "Matched Family (email)", "Notes"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

withdrawals_tab_rows.sort(key=lambda w: w["date"])
for i, w in enumerate(withdrawals_tab_rows):
    rownum = 3 + i
    updated_install_val = None
    if w["updated_install_raw"]:
        try:
            updated_install_val = float(w["updated_install_raw"].replace("$", "").replace(",", ""))
        except ValueError:
            updated_install_val = w["updated_install_raw"]  # leave as text rather than erroring on something unparseable
    vals = [w["date"].strftime("%Y-%m-%d"), w["parent_name"], w["participant_name"], EVENT_LABEL[w["signup"]],
            w["refund_type_raw"].title(), updated_install_val, w["refund_amt_expected"], w["credit_amt_found"],
            w["matched_email"] or "(no match)", w["note"]]
    body_row(ws, rownum, vals, len(vals), start_col=1)
    for col in (6, 7, 8):
        ws.cell(row=rownum, column=col).number_format = CURRENCY
    ws.cell(row=rownum, column=10).alignment = Alignment(vertical="center", wrap_text=True)
    if not w["matched_email"]:
        for c in range(1, 11):
            ws.cell(row=rownum, column=c).fill = ORANGE_FILL

last_data_row = 2 + len(withdrawals_tab_rows)
# Kept under its own name, like SIGNUPS_LAST_ROW/TX_LAST_ROW -- the Dashboard's new Refunds section (added Oct 15,
# 2026, per Michal) references this directly via a live SUM formula, so it must stay a valid range even with zero
# withdrawal rows (row 3 then just holds "No withdrawals logged yet." text, which SUM() silently ignores).
WITHDRAWALS_LAST_ROW = max(last_data_row, 3)
ws.freeze_panes = "A3"
autofit(ws, [12, 20, 22, 16, 22, 20, 16, 18, 26, 50])
ws.auto_filter.ref = f"A2:J{max(last_data_row, 2)}"

if withdrawals_tab_rows:
    trow = last_data_row + 2
    n_refund_confirmed = sum(1 for w in withdrawals_tab_rows if w["credit_amt_found"])
    total_confirmed = sum(w["credit_amt_found"] or 0 for w in withdrawals_tab_rows)
    ws.cell(row=trow, column=1,
            value=f"{len(withdrawals_tab_rows)} withdrawal(s) logged -- {n_refund_confirmed} with a confirmed Authorize.net credit").font = SECTION_FONT
    ws.cell(row=trow, column=8, value=total_confirmed).number_format = CURRENCY
    ws.cell(row=trow, column=8).font = BOLD_FONT
else:
    ws.cell(row=3, column=1, value="No withdrawals logged yet.").font = SUB_FONT

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("withdrawal tab rows:", len(withdrawals_tab_rows))

# =========================================================================
# SHEET: Donations
# =========================================================================
def _roster_identity(key, event):
    """(parent name, roster email) for whichever Roster / Seminar Roster parent this canonical family email
    belongs to -- used so the Donations tab credits the parent as recorded on the roster, not just whichever name/
    email happened to be printed on the receipt (the same email may be a work address, per the Sharon Blumenkranz
    case, or the payment may come from either parent)."""
    rows = seminar_roster_rows if event == SEMINAR else roster_rows
    pairs = [("Parent 1 Email", "Parent 1 Contact Name"), ("Parent 2 Email", "Parent 2 Contact Name")]
    for r in rows:
        for email_col, name_col in pairs:
            e = r[email_col].strip().lower().replace("\\_", "_")
            if e == key:
                return (r[name_col].strip() or key), key
    return key, key  # shouldn't happen for an auto-detected donation, but fall back safely rather than erroring

donation_rows = []
for row in tx_rows:
    txn_id = row[8]
    amt = donations_map.get(txn_id)
    if not amt:
        continue
    signup = row[2]
    lookup_key = resolved_email.get(txn_id, row[5])
    p_name, p_email = _roster_identity(lookup_key, signup)
    donation_rows.append([p_name, p_email, EVENT_LABEL.get(signup, signup), amt, row[0], row[9]])
donation_rows.sort(key=lambda r: r[4])

ws = wb.create_sheet("Donations")
ws["A1"] = ("Every Donation Add-on currently logged on the Transactions tab (column M) -- that's the auto-detected "
            "embedded-donation Seminar cases plus (added Sep 4, 2026, per Michal) any storefront order whose item "
            "is named \"$X Donation\", auto-detected by that item name; anything logged there by hand in the "
            "future will need to be added here too on the next refresh. Parent Name and Email come from the "
            "Roster / Seminar Roster, not the raw payer name/email on the receipt, so this reflects who actually "
            "gets credit for the donation regardless of which parent's card was charged -- a storefront donation "
            "that can't be matched to anyone on either roster falls back to showing the buyer's own name/email "
            "instead (see Read Me).")
ws["A1"].font = SUB_FONT
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A1:F1")
ws.row_dimensions[1].height = 42

headers = ["Parent Name (Roster)", "Email (Roster)", "Signup", "Donation Amount", "Date", "Kid(s) Covered"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

for i, r in enumerate(donation_rows):
    rownum = 3 + i
    body_row(ws, rownum, r, len(r), start_col=1)
    ws.cell(row=rownum, column=4).number_format = CURRENCY

last_row = 2 + len(donation_rows)
if donation_rows:
    ws.cell(row=last_row + 2, column=1, value="Total").font = BOLD_FONT
    tot_cell = ws.cell(row=last_row + 2, column=4, value=f"=SUM(D3:D{last_row})")
    tot_cell.font = BOLD_FONT
    tot_cell.number_format = CURRENCY
else:
    ws.cell(row=3, column=1, value="No donations logged yet.").font = SUB_FONT

ws.freeze_panes = "A3"
autofit(ws, [22, 26, 20, 16, 12, 34])
ws.auto_filter.ref = f"A2:F{max(last_row, 2)}"

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("donation rows written:", len(donation_rows))

# =========================================================================
# SHEET: Store Purchases
# =========================================================================
# Tapuach Store merchandise charges -- logged on Transactions for a complete ledger, but excluded from Signups and
# from the Dashboard's by-signup-type totals since they aren't Annual/Seminar signup fees. Pulled out onto their
# own tab (Michal, Aug 28, 2026) so they're easy to see together instead of scattered among signup payments.
store_rows = [row for row in tx_rows if row[2] == STORE]

ws = wb.create_sheet("Store Purchases")
ws["A1"] = ("Every Tapuach Store (merchandise) charge found in the receipt emails -- these are excluded from Signups and from the "
            "Dashboard's Annual/Seminar totals since they aren't signup fees, but are totaled separately on the Dashboard. Item, "
            "Participant, and Grade (Participant/Grade added Sep 4, 2026, per Michal) come from the new \"Shevet Tapuach\" "
            "storefront's own order-confirmation emails -- Authorize.net's older Tapuach Store receipts don't carry that level of "
            "detail, so those rows show all three blank. Grade is the \"Grade: <N> Grade\" value from the order email itself "
            "(2nd-12th grade), taken as-is rather than looked up on a roster, so it's only blank when the order wasn't placed "
            "\"for\" a specific participant. A row whose Item is a \"$X Donation\" (also added Sep 4, 2026) is a donation add-on "
            "sold through the same storefront checkout, not merchandise -- it's still shown here since that's where the payment "
            "came through, but is also broken out on the Donations tab (see its Notes column and the Read Me).")
ws["A1"].font = SUB_FONT
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A1:J1")
ws.row_dimensions[1].height = 28

headers = ["Date", "Payer Name", "Payer Email", "Amount", "Card", "Invoice #", "Item", "Participant", "Grade", "Notes"]
for i, h in enumerate(headers):
    ws.cell(row=2, column=1+i, value=h)
style_header(ws, 2, len(headers), start_col=1)

for i, row in enumerate(store_rows):
    rownum = 3 + i
    _item_disp, _participant_disp, _grade_disp = store_item_info.get(row[8], ("", "", ""))
    vals = [row[0], row[4], row[5], row[3], row[6], row[7], _item_disp, _participant_disp, _grade_disp, row[11]]
    body_row(ws, rownum, vals, len(vals), start_col=1)
    ws.cell(row=rownum, column=4).number_format = CURRENCY

last_row = 2 + len(store_rows)
if store_rows:
    ws.cell(row=last_row + 2, column=1, value="Total").font = BOLD_FONT
    tot_cell = ws.cell(row=last_row + 2, column=4, value=f"=SUM(D3:D{last_row})")
    tot_cell.font = BOLD_FONT
    tot_cell.number_format = CURRENCY
else:
    ws.cell(row=3, column=1, value="No Tapuach Store purchases logged yet.").font = SUB_FONT

ws.freeze_panes = "A3"
autofit(ws, [11, 20, 26, 12, 12, 15, 22, 20, 14, 40])
ws.auto_filter.ref = f"A2:J{max(last_row, 2)}"

wb.save("/home/claude/tapuach_report/scaffold.xlsx")
print("store purchase rows written:", len(store_rows))

# =========================================================================
# SHEET 6: Dashboard
# =========================================================================
ws = wb.create_sheet("Dashboard")
ws["B2"] = "Payments Dashboard"
ws["B2"].font = TITLE_FONT
ws["B3"] = "Live totals pulled from the Transactions and Signups tabs."
ws["B3"].font = SUB_FONT

def stat_box(ws, row, label, formula, fmt=None, fill=None):
    ws.cell(row=row, column=2, value=label).font = BOLD_FONT
    c = ws.cell(row=row, column=5, value=formula)
    c.font = Font(name=FONT, bold=True, size=12, color="1F4E78")
    c.alignment = Alignment(horizontal="right")
    if fmt:
        c.number_format = fmt
    for col in (2, 3, 4, 5):
        ws.cell(row=row, column=col).border = BORDER
    if fill:
        for col in (2, 3, 4, 5):
            ws.cell(row=row, column=col).fill = fill
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)

ws["B5"] = "Overall"
ws["B5"].font = SECTION_FONT
stat_box(ws, 6, "Total collected to date (all signups, all transactions) -- original payment amount",
         f"=SUM(Transactions!$D$3:$D${TX_LAST_ROW})", CURRENCY)
stat_box(ws, 7, "Of which, Donation Add-ons (logged in Transactions column M)",
         f"=SUM(Transactions!$M$3:$M${TX_LAST_ROW})", CURRENCY)
stat_box(ws, 8, "Authorize.net fees (~2% of charges, none on refunds -- added Aug 30, 2026)",
         f"=SUM(Transactions!$O$3:$O${TX_LAST_ROW})", CURRENCY, ORANGE_FILL)
stat_box(ws, 9, "Total collected after fees (what actually lands in the account)",
         f"=SUM(Transactions!$P$3:$P${TX_LAST_ROW})", CURRENCY, GREEN_FILL)
stat_box(ws, 10, "Total expected (sum of Expected Total entered on Signups)",
         f"=SUM(Signups!$I$3:$I${SIGNUPS_LAST_ROW})", CURRENCY)
stat_box(ws, 11, "Total outstanding / still owed (only where Expected Total is filled in)",
         f"=SUM(Signups!$J$3:$J${SIGNUPS_LAST_ROW})", CURRENCY)
stat_box(ws, 12, "Transactions logged", f"=COUNTIF(Transactions!$A$3:$A${TX_LAST_ROW},\"<>\")", "0")
stat_box(ws, 13, "Rows flagged for review on Transactions (roster match issues)",
         f"=COUNTIFS(Transactions!$A$3:$A${TX_LAST_ROW},\"<>\",Transactions!$K$3:$K${TX_LAST_ROW},\"<>OK\")", "0", ORANGE_FILL)

ws["B15"] = "By signup status"
ws["B15"].font = SECTION_FONT
stat_box(ws, 16, "Signups paid in full", f'=COUNTIF(Signups!$K$3:$K${SIGNUPS_LAST_ROW},"Paid in full")', "0", GREEN_FILL)
stat_box(ws, 17, "Signups mid installment-plan", f'=COUNTIF(Signups!$K$3:$K${SIGNUPS_LAST_ROW},"In progress (installment plan)")', "0", YELLOW_FILL)
stat_box(ws, 18, "Signups registered but not yet paid", f'=COUNTIF(Signups!$K$3:$K${SIGNUPS_LAST_ROW},"Not yet paid")', "0", RED_FILL)
stat_box(ws, 19, "Signups still needing an Expected Total entered", f'=COUNTIF(Signups!$K$3:$K${SIGNUPS_LAST_ROW},"Enter expected total")', "0", ORANGE_FILL)
stat_box(ws, 20, "Signups with a withdrawal (refund/adjust -- see Withdrawals tab)",
         f'=COUNTIF(Signups!$K$3:$K${SIGNUPS_LAST_ROW},"Refund (withdrawal)")'
         f'+COUNTIF(Signups!$K$3:$K${SIGNUPS_LAST_ROW},"Adjusted (withdrawal)")'
         f'+COUNTIF(Signups!$K$3:$K${SIGNUPS_LAST_ROW},"Refund + Adjusted (withdrawal)")', "0", PURPLE_FILL)

ws["B21"] = "By signup type -- original payment / after Authorize.net fee"
ws["B21"].font = SECTION_FONT
stat_box(ws, 22, "Annual Membership -- collected so far (original)",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{ANNUAL}",Transactions!$D$3:$D${TX_LAST_ROW})', CURRENCY)
stat_box(ws, 23, "Annual Membership -- after fee",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{ANNUAL}",Transactions!$P$3:$P${TX_LAST_ROW})', CURRENCY)
stat_box(ws, 24, "Opening Seminar -- collected so far (original)",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{SEMINAR}",Transactions!$D$3:$D${TX_LAST_ROW})', CURRENCY)
stat_box(ws, 25, "Opening Seminar -- after fee",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{SEMINAR}",Transactions!$P$3:$P${TX_LAST_ROW})', CURRENCY)
stat_box(ws, 26, "Yom Tzofe -- collected so far (original)",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{YOM_TZOFE}",Transactions!$D$3:$D${TX_LAST_ROW})', CURRENCY)
stat_box(ws, 27, "Yom Tzofe -- after fee",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{YOM_TZOFE}",Transactions!$P$3:$P${TX_LAST_ROW})', CURRENCY)
stat_box(ws, 28, "Tapuach Store -- collected so far (see Store Purchases tab, original)",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{STORE}",Transactions!$D$3:$D${TX_LAST_ROW})', CURRENCY)
stat_box(ws, 29, "Tapuach Store -- after fee",
         f'=SUMIF(Transactions!$C$3:$C${TX_LAST_ROW},"{STORE}",Transactions!$P$3:$P${TX_LAST_ROW})', CURRENCY)

# Added Oct 15, 2026, per Michal ("can you add the refunds to the dashboard"). Row 32 is the authoritative total --
# every Authorize.net Credit (refund) transaction is logged on the Transactions tab as a negative Amount (see
# Transactions-sheet build above), so summing the negative side of column D and flipping the sign gives the true
# total refunded to date regardless of whether it's been matched to a specific logged withdrawal yet. Row 33 is a
# narrower cross-check: it sums the Withdrawals tab's own "Refund Amount" column (H) over WITHDRAWALS_LAST_ROW, i.e.
# only withdrawals Michal has actually logged with a confirmed Authorize.net credit -- it can be smaller than row 32
# if a refund was issued but hasn't been logged as a withdrawal yet (or larger would indicate a data problem). Row 34
# is just a count, for a quick sense of volume.
ws["B31"] = "Refunds (Authorize.net credits)"
ws["B31"].font = SECTION_FONT
stat_box(ws, 32, "Total refunded to date (all Authorize.net credits on Transactions)",
         f'=-SUMIF(Transactions!$D$3:$D${TX_LAST_ROW},"<0")', CURRENCY)
stat_box(ws, 33, "Of which, confirmed via a logged withdrawal (see Withdrawals tab)",
         f'=SUM(Withdrawals!$H$3:$H${WITHDRAWALS_LAST_ROW})', CURRENCY)
stat_box(ws, 34, "Refund (credit) transactions logged",
         f'=COUNTIF(Transactions!$D$3:$D${TX_LAST_ROW},"<0")', "0")

ws["B36"] = ("Note: the $1.00 verification charge from Aug 10 is included in \"Total collected\" (it's negligible) but excluded from "
             "the by-signup-type totals' underlying pricing assumptions. \"Total collected\" and the by-signup-type totals include "
             "Donation Add-ons; Signups' Total Paid and this Dashboard's Total expected/outstanding do not. Tapuach Store purchases are "
             "included in \"Total collected\" but not in Total expected/outstanding, since they're not a signup fee. \"Not yet paid\" "
             "rows come from cross-checking the Seminar Roster / Yom Tzofe Roster tabs against Transactions -- the same check hasn't "
             "been run for Annual Membership yet. Authorize.net fees (added Aug 30, 2026, per Michal) are computed as 2% of each "
             "transaction's gross Amount on the Transactions tab (column O), with 0 fee on refunds/credits; \"Total expected\"/\"Total "
             "outstanding\" and Signups' Total Paid are unaffected by fees -- they still track the gross signup price, since that's "
             "what a family owes regardless of what the processor keeps. Refunds section (added Oct 15, 2026, per Michal) is also the "
             "reason the Payouts and Payout by Event tabs now subtract that day's refunds from Payments Hub's gross \"Total Batch "
             "Volume\" before comparing it to Authorize.net's settled totals, which are already net of refunds -- see those tabs' own "
             "notes. See Read Me for details.")
ws["B36"].font = SUB_FONT
ws["B36"].alignment = Alignment(wrap_text=True)
ws.merge_cells("B36:F39")
ws.row_dimensions[36].height = 110

autofit(ws, [3, 55, 3, 3, 16])
ws.sheet_view.showGridLines = False

# Reorder sheets: Read Me, Dashboard, Signups, Transactions, Roster, Price Reference
order = ["Read Me", "Dashboard", "Signups", "Scholarships", "Withdrawals", "Donations", "Store Purchases", "Transactions", "Reconciliation", "Payouts", "Payout by Event", "Venmo", "Roster", "Seminar Roster", "Yom Tzofe Roster", "Price Reference"]
wb._sheets = [wb[name] for name in order]
wb.active = 0

wb.save("/home/claude/tapuach_report/Tapuach_Payments_Tracker.xlsx")
print("DONE")
