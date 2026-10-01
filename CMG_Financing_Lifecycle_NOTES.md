# CMG Financing Lifecycle — working notes

Handover notes so work can resume where it stopped.

- **Current version:** `CMG_Financing_Lifecycle_v8.html` (rough cut, 1 Oct 2026)
- **Earlier versions:** `archive/cmg-financing-lifecycle/` (v3 SVG; v4 to v7 HTML). v6 is the last version before the FI-led model; v7 was a same-day draft with two loops and a pop-up.
- **Project:** World Bank PAMOJA (P178813), Sub-Component 1.2 — capital injection to Community Microfinance Groups (CMGs)

## What it is

A one-page interactive graphic, with no scrolling, that shows how money and obligations move between PO-RALG, the council, the FI, the CMG and the woman entrepreneur. Open it in a browser and click a step or a label to see what applies.

v8 shows three loans, one inside the next, and peels one loop at a time:

- **Loop 1, PO-RALG → FI:** on-lending agreement, funding, reporting, repayment into the Pamoja fund, loss sharing. Opens first.
- **Loop 2, FI → CMG:** the deck's seven stages (Apply to If not repaid).
- **Loop 3, CMG → member:** the four lifecycle steps kept from v6.

How to move between loops: three numbered tabs, in the loop colours, sit on the top shelf of the box (1 PO-RALG · sets, funds and oversees; 2 FI · lends to CMGs; 3 CMG · lends to women members). The current loop's tab is filled. A dashed cue card inside loops 1 and 2 also opens the next loop in, and Esc steps back out. A small "full picture" link at the right of the shelf shows all three loops stacked; this is the only view that scrolls.

## Sources

**A. Financing Agreement extracts** (as in v6)

1. **COM clause.** PO-RALG prepares and adopts the Community Operations Manual, covering areas (a)–(g): structure, entry, lending terms, money flow, use of funds, accountability, safeguards.
2. **CMG obligations, item c (i)–(vii).** Lend only to women members; manage the funds with diligence (Anti-Corruption Guidelines); monitor; keep accounts; allow inspection; provide information; refund unused funds within 60 days of the Closing Date or a wind-up decision.
3. **Paragraph 6.** The Microfinancing Agreement between the CMG and the entrepreneur: amount, currency, charges and schedule; the CMG's remedies; the entrepreneur's obligations (c)(i)–(iv) and (viii). **Numbering jumps from iv to viii; still to check against the signed agreement.**
4. **Paragraph 7.** PO-RALG ensures the CMG enforces its rights. The CMG may not assign, amend or waive the standard terms.

**B. Deck "CMG Lending: the FI-led model", draft v0.4, 1 Oct 2026** (new in v7). Based on the English translation of the CMG Lending Programme Guideline. Source of loops 1 and 2, the council and FI roles, and the "still to settle" points. The deck says TAMISEMI; this graphic says PO-RALG (Anand will change the deck to match).

## Design decisions (agreed with Anand)

- Visual = **Option 3**: a U-shaped lifecycle of step cards, coloured by layer, with a nested-layers legend (Option C) + lifecycle (Option D). The look is unchanged since v6.
- **v7, 1 Oct 2026: the FI loan replaces the Deposit Agreement.** v6 steps 1, 2 and 7 (CMG selected, Deposit Agreement, Close-out & refund) are replaced by the deck's stages.
- **v8, 1 Oct 2026: three loops, peeled one at a time** ("peel the onion"), with a low-key full picture. Steps are numbered loop.step.
- **Loop 1 steps:** 1.1 On-lending agreement → 1.2 Fund the FI → [FI lends to CMGs] → 1.3 Report → 1.4 Repay → 1.5 Share the loss.
- **Loop 2 steps:** 2.1 Apply → 2.2 Check → 2.3 Decide → 2.4 Contract → 2.5 Pay out → [CMG lends to women members] → 2.6 Repay → 2.7 If not repaid.
- **Loop 3 steps** (v6 content unchanged): 3.1 Subproject selected → 3.2 Microfinancing Agreement → 3.3 Lend · Monitor · Report → 3.4 Repayment.
- **Frames (the onion):** each lender has a frame in its colour: PO-RALG (always shown), then FI, then CMG, nested as you go in. The frames carry no labels of their own; the shelf tabs name them.
- **Parties:** five coloured parties — PO-RALG, Council, FI, CMG, Entrepreneur. The council is not a lender, so it has no loop: it pays the FI (1.2) and checks and shortlists (2.2, 2.3). The Pamoja fund is not a party; it appears only as where the money goes (TZS chip on 1.4).
- **Revolving loops:** in loop 1, repayments return to the Pamoja fund (1.4 → 1.2), labelled "relend or loss? still to settle". In loop 3, the CMG re-lends repayments (3.4 → 3.1, new Subproject and new agreement).
- **If not repaid:** red dashed arrows (1.4 → 1.5 and 2.6 → 2.7).
- **Shared steps:** side tabs mark the two parties to an agreement (1.1, 2.4, 3.2). Badges mark who carries out the step.
- Symbols: TZS chip = money moves; ◆ = obligations start, apply or end; ⚖ = enforcement; dashed ↻ = revolving funds; dashed card = cue for the next loop in.
- Colours: PO-RALG `#1f4e79`, Council `#4a8fd1`, FI `#b8437f`, CMG `#1f9e89`, Entrepreneur `#e39b2d`, money `#2e7d32`.
- **PO-RALG stands out.** A dark-blue frame surrounds the whole diagram ("sets, funds and oversees"). Its labels have a blue outline, with normal text weight.
- **Label rows:** "COM covers" (a–g) at the top; "PO-RALG must", "Council must", "FI must", "CMG must" and "Entrepreneur must" at the bottom. Clicking one highlights the steps it applies to. If it applies in another loop, the cue card lights up and that loop's tab is underlined. The reverse also works: pointing at or clicking a step lights up, in the top and bottom rows, every label the side panel lists for it. A clicked step gets a heavy outline while the rest of the loop and the other labels fade. Detail appears only in the side panel, so nothing is shown twice.
- **Still to settle:** the deck's open points appear only in the side panel of the step they affect.
- **Emphasis (v8):** the loop box is the main thing on the page: white, large, with a soft shadow. The five "must" rows are small. Step cards no longer carry the COM letters (a–g); the COM row and the side panel still show which areas apply.

## Style rules

- **Do not show the World Bank.** Use **PO-RALG**, not "Government" or TAMISEMI.
- Currency: **TZS**, not TSh.
- Plain labels: "Lending terms" (not "Money terms"), "Allow inspection" (not "Open up").
- Minimalist look; one page with no scrolling at 1280×720 or larger.
- Keep strict version numbers (vN). Every deliverable has a glossary.

## Mappings that are my interpretation (to be validated)

- COM areas → steps: a all; b 2.1, 2.2, 2.3; c 2.4, 3.2; d 1.1, 1.2, 2.4, 2.5; e 3.1; f 3.3; g 3.1, 3.3.
- COM area d was written for the Deposit Agreement. Its detail is reworded as "flow of funds to the CMG, now through the council and the FI; agreement templates".
- "Council must" and "FI must" come from the deck (a proposed model), not from the Financing Agreement. The side panel labels them "role · proposed FI-led model".
- **Loop 1 is assembled from pieces of the deck**, which has no separate PO-RALG–FI process. Splitting the deck's stages: Pay out became 1.2 (council pays FI) and 2.5 (FI pays CMG); Repay became 2.6 (CMG to FI) and 1.4 (FI to Pamoja fund); If not repaid became 2.7 (FI recovers) and 1.5 (loss split). "Report" (1.3) is shown as a step although reporting is continuous.
- 1.5 "Share the loss" carries PO-RALG and FI badges. The deck only says the loss is "split by an agreed share" and does not name the parties.
- "FI obligations start" on 1.1 is inferred.
- PO-RALG row: "Select CMGs" is removed (the council shortlists, the FI decides). "Sign & fund" now means the on-lending agreement and allocating funds to the council. "Redeploy" becomes "Receive repayments". "Share losses" is added.
- CMG row: "Return" (refund unused funds within 60 days) is replaced by "Repay" (repay the FI). "Apply" is added. "CMG obligations start" sits on 2.4 and "end" on 2.6.
- The CMG re-lending repayments while it still owes the FI is carried over from v6; the deck does not cover it.

## Known constraints

- Built for landscape screens of 1280×720 or larger; not for phones or portrait screens.
- Present from a browser; interactivity is lost if it is pasted into PowerPoint.
- Click or tap works everywhere; hover works only with a mouse.
- The panel text must stay short (no scrolling).
- Printing shows only the loop on screen.
- The full picture scrolls inside the frame; it is for reference, not for presenting.

## Open items / possible next steps

- Check the jump from (iv) to (viii) in paragraph 6(c).
- Confirm "M-ESCP" (probably Mainland ESCP).
- Validate the mappings above.
- Check with the legal text whether the Deposit Agreement obligations (refund within 60 days, Closing Date) still apply once the FI is the lender.
- From the deck, still to settle: who lends to the FI; the loss split; PO-RALG's role in repayment; what happens to money in the Pamoja fund; one transfer or instalments; a definition of default; a time limit and appeal for the FI's decision; the repeated checks.
- Possible v9: a full graphic polish; a stable file name for sharing.

## Glossary

| Term | Meaning |
|---|---|
| COM | Community Operations Manual |
| CMG | Community Microfinance Group |
| PO-RALG | President's Office – Regional Administration and Local Government (TAMISEMI) |
| Council | The Local Government Authority (LGA): ward, CMT and Director |
| FI | Financial institution: the partner bank that lends to the CMG |
| CMT | Council Management Team |
| CDO / CD | Community Development Officer / community development |
| DCRC | District Credit Review Committee |
| DIT | District Implementation Team |
| Wezesha | PO-RALG's portal for group registration and loans |
| GePG | Government e-Payment Gateway, which issues control numbers |
| Pamoja fund | The pool at Bank of Tanzania that repayments return to |
| Account B | The council's collection account, which repayments pass through on the way to the Pamoja fund |
| On-lending agreement | Agreement under which funds are lent to the FI, which lends them on to CMGs |
| Loan contract | Contract between the FI and a CMG (step 2.4) |
| Loop | One loan and its repayment: PO-RALG → FI, FI → CMG, or CMG → member |
| M-ESCP | Environmental and Social Commitment Plan ("M" probably Mainland) |
| TZS | Tanzanian shilling |
| Deposit Agreement | v6 only: contract between PO-RALG and a CMG for the capital injection. Replaced in v7 by the FI loan contract |
| Microfinancing Agreement | Contract between the CMG and a woman entrepreneur for her loan |
| Subproject | A woman entrepreneur's productive investment financed by the CMG |
| Shortlist | A list of CMGs the council puts forward, which the FI then decides on |
