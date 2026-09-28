# CMG Financing Lifecycle — working notes

Handover notes so work can resume where it stopped.

- **Current version:** `CMG_Financing_Lifecycle_v6.html` (rough cut, 28 Sep 2026)
- **Earlier versions:** `archive/cmg-financing-lifecycle/` (v3 SVG, v4 and v5 HTML)
- **Project:** World Bank PAMOJA (P178813), Sub-Component 1.2 — capital injection to Community Microfinance Groups (CMGs)

## What it is

A one-page interactive graphic, with no scrolling, that shows how money and obligations move between PO-RALG, the CMG and the woman entrepreneur. Open it in a browser and click a step or a label to see what applies.

## Source texts it is based on (Financing Agreement extracts)

1. **COM clause.** PO-RALG prepares and adopts the Community Operations Manual, covering areas (a)–(g): structure, entry, lending terms, money flow, use of funds, accountability, safeguards.
2. **CMG obligations, item c (i)–(vii).** Lend only to women members; manage the funds with diligence (Anti-Corruption Guidelines); monitor; keep accounts; allow inspection; provide information; refund unused funds within 60 days of the Closing Date or a wind-up decision.
3. **Paragraph 6.** The Microfinancing Agreement between the CMG and the entrepreneur: amount, currency, charges and schedule; the CMG's remedies; the entrepreneur's obligations (c)(i)–(iv) and (viii). **Numbering jumps from iv to viii; still to check against the signed agreement.**
4. **Paragraph 7.** PO-RALG ensures the CMG enforces its rights. The CMG may not assign, amend or waive the standard terms.

## Design decisions (agreed with Anand)

- Visual = **Option 3**: a U-shaped lifecycle of 7 steps, coloured by layer, with a nested-layers legend (Option C) + lifecycle (Option D).
- Steps: 1 CMG selected → 2 Deposit Agreement → 3 Subproject selected → 4 Microfinancing Agreement → 5 Lend · Monitor · Report → 6 Repayment → 7 Close-out & refund.
- **Revolving loops:** the CMG re-lends repayments (6 → 3, which needs a new Subproject and a new agreement). PO-RALG redeploys refunds **always to a new CMG** (7 → 1). Refunded funds **can** be redeployed.
- **Shared steps:** side tabs mark the two parties to an agreement (steps 2 and 4). Badges mark who carries out the step (5 has two badges).
- Symbols: TZS chip = money moves; ◆ = obligations start, apply or end; ⚖ = enforcement; dashed ↻ = revolving funds.
- Colours: PO-RALG `#1f4e79`, CMG `#1f9e89`, Entrepreneur `#e39b2d`, money `#2e7d32`.
- **PO-RALG stands out.** A dark-blue frame surrounds the whole diagram ("sets, funds and oversees"). Its labels have a blue outline, with normal text weight.
- **Label rows:** "COM covers" (a–g) at the top; "PO-RALG must", "CMG must" and "Entrepreneur must" at the bottom. Clicking one highlights the steps it applies to. Detail appears only in the side panel, so nothing is shown twice.

## Style rules

- **Do not show the World Bank.** Use **PO-RALG**, not "Government".
- Currency: **TZS**, not TSh.
- Plain labels: "Lending terms" (not "Money terms"), "Allow inspection" (not "Open up").
- Minimalist look; one page with no scrolling at 1280×720 or larger.
- Keep strict version numbers (vN). Every deliverable has a glossary.

## Mappings that are my interpretation (to be validated)

- COM areas → steps: a all; b 1; c 2, 4; d 2; e 3; f 5; g 3, 5.
- PO-RALG "Select CMGs" and "Oversee" are inferred. The clauses don't say who selects CMGs, and "Oversee" draws on the CMG reporting to DITs and the inspection and information rights. "Redeploy" comes from Anand's instruction.

## Known constraints

- Built for landscape screens of 1280×720 or larger; not for phones or portrait screens.
- Present from a browser; interactivity is lost if it is pasted into PowerPoint.
- Click or tap works everywhere; hover works only with a mouse.
- The panel text must stay short (no scrolling).
- Printing shows only the default view.

## Open items / possible next steps

- Check the jump from (iv) to (viii) in paragraph 6(c).
- Confirm "M-ESCP" (probably Mainland ESCP).
- Validate the mappings above.
- Possible v7: a full graphic polish; a stable file name for sharing.

## Glossary

| Term | Meaning |
|---|---|
| COM | Community Operations Manual |
| CMG | Community Microfinance Group |
| PO-RALG | President's Office – Regional Administration and Local Government (TAMISEMI) |
| DIT | District Implementation Team |
| M-ESCP | Environmental and Social Commitment Plan ("M" probably Mainland) |
| TZS | Tanzanian shilling |
| Deposit Agreement | Contract between PO-RALG and a CMG for the capital injection |
| Microfinancing Agreement | Contract between the CMG and a woman entrepreneur for her loan |
| Subproject | A woman entrepreneur's productive investment financed by the CMG |
| Wind-up | A CMG's decision to close down or leave the project |
