# Legal documents

## Teper–KGK Trading Card Joint Venture Operating Agreement

`Teper-KGK-Trading-Card-JV-Operating-Agreement.docx` is a **draft for attorney review** of a
joint venture / operating agreement between Teper Holdings and Group KGK Inc. covering the joint
purchase, custody, and sale of graded trading cards.

Key structure:

- Teper Holdings funded 100% of the acquisition (the Initial Contribution).
- Group KGK Inc. owes a Deferred Contribution (its 50% share) by a deadline; until paid it holds a
  Contribution Right rather than a percentage interest. An alternative 50/50-from-day-one structure
  is flagged in a drafting note for the attorney to choose.
- Unanimous consent for Major Decisions (sale, pricing, grading, encumbrance), Teper Holdings as
  Managing Member, distribution waterfall, right of first offer, deadlock buy-sell, transfer
  restrictions, tax provisions, and dissolution.
- Schedules: A (the Cards), B (capital and percentage interests), C (payment schedule and expense
  ledger), plus an Attorney Review Memorandum listing every open item.

All placeholders are in `[brackets]` and highlighted yellow in the Word file.

### Regenerating

The document is generated from `build_agreement.js` using the `docx` npm package:

```bash
npm install docx jszip
node legal/build_agreement.js
```
