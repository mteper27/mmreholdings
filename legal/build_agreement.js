// Generates the Teper Holdings / Group KGK Inc. trading card joint venture operating agreement.
// Run:  node legal/build_agreement.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Header, Footer,
  PageNumber, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  LevelFormat, TabStopType, PageBreak,
} = require("docx");

const FONT = "Times New Roman";
const BODY = 22; // 11pt in half-points
const OUT = path.join(__dirname, "Teper-KGK-Trading-Card-JV-Operating-Agreement.docx");

// ---- helpers ---------------------------------------------------------------
// Splits text on [bracketed placeholders] and highlights them yellow so the
// reviewing attorney can find every open item at a glance.
function runs(text, opts = {}) {
  // Walk the string tracking bracket depth so nested placeholders such as
  // "[at [Address] / in a vault]" are highlighted as one unit.
  const out = [];
  const push = (t, hl) => { if (t) out.push(new TextRun({ text: t, font: FONT, size: BODY, ...(hl ? { highlight: "yellow" } : {}), ...opts })); };
  let depth = 0, buf = "";
  for (const ch of text) {
    if (ch === "[") { if (depth === 0) { push(buf, false); buf = ""; } depth++; buf += ch; }
    else if (ch === "]" && depth > 0) { depth--; buf += ch; if (depth === 0) { push(buf, true); buf = ""; } }
    else buf += ch;
  }
  push(buf, depth > 0);
  return out;
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 160, line: 276 },
    alignment: AlignmentType.JUSTIFIED,
    ...opts.para,
    children: runs(text, opts.run),
  });
}

// "1.1  Heading. Body text" -> bold number + bold caption, plain body.
function sec(num, caption, body) {
  return new Paragraph({
    spacing: { after: 160, line: 276 },
    alignment: AlignmentType.JUSTIFIED,
    children: [
      new TextRun({ text: `${num}\t`, font: FONT, size: BODY, bold: true }),
      new TextRun({ text: `${caption}. `, font: FONT, size: BODY, bold: true }),
      ...runs(body),
    ],
    tabStops: [{ type: TabStopType.LEFT, position: 720 }],
  });
}

function sub(letter, text) {
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: 1080, hanging: 360 },
    children: [new TextRun({ text: `(${letter})\t`, font: FONT, size: BODY }), ...runs(text)],
    tabStops: [{ type: TabStopType.LEFT, position: 1080 }],
  });
}

function article(num, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 200 },
    keepNext: true,
    children: [new TextRun({ text: `ARTICLE ${num}`, font: FONT, size: 24, bold: true }),
               new TextRun({ break: 1, text: title.toUpperCase(), font: FONT, size: 24, bold: true })],
  });
}

function centered(text, size = BODY, bold = true, after = 120) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after }, children: runs(text, { bold, size }) });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 100, line: 276 },
    alignment: AlignmentType.LEFT,
    children: runs(text),
  });
}

function blank(after = 120) { return new Paragraph({ spacing: { after }, children: [] }); }

function note(text) {
  return new Paragraph({
    spacing: { after: 160, line: 276 },
    alignment: AlignmentType.JUSTIFIED,
    shading: { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" },
    indent: { left: 360, right: 360 },
    children: [new TextRun({ text: "Drafting Note: ", font: FONT, size: 20, bold: true, italics: true }),
               ...runs(text, { size: 20, italics: true })],
  });
}

const cellBorders = { top: { style: BorderStyle.SINGLE, size: 4, color: "999999" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
                      left: { style: BorderStyle.SINGLE, size: 4, color: "999999" }, right: { style: BorderStyle.SINGLE, size: 4, color: "999999" } };

function table(headers, rows, widths) {
  const mk = (t, head) => new TableCell({
    width: { size: 0, type: WidthType.DXA },
    borders: cellBorders,
    shading: head ? { type: ShadingType.CLEAR, fill: "D9D9D9", color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ spacing: { after: 0 }, children: runs(t, { bold: head, size: 20 }) })],
  });
  const fix = (row, head) => new TableRow({ tableHeader: head, children: row.map((t, i) => { const c = mk(t, head); c.options.width = { size: widths[i], type: WidthType.DXA }; return c; }) });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [fix(headers, true), ...rows.map(r => fix(r, false))],
  });
}

function sigBlock(entity, roleLine) {
  const line = (label) => new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: `${label}\t`, font: FONT, size: BODY }), new TextRun({ text: "", font: FONT, size: BODY })],
    tabStops: [{ type: TabStopType.LEFT, position: 1440 }, { type: TabStopType.LEFT, position: 6480, leader: "underscore" }] });
  return [
    new Paragraph({ spacing: { before: 240, after: 120 }, keepNext: true, children: runs(entity, { bold: true }) }),
    para(roleLine, { para: { alignment: AlignmentType.LEFT } }),
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "By:\t", font: FONT, size: BODY }), new TextRun({ text: "\t", font: FONT, size: BODY })],
      tabStops: [{ type: TabStopType.LEFT, position: 1440 }, { type: TabStopType.LEFT, position: 6480, leader: "underscore" }] }),
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Name:\t", font: FONT, size: BODY }), new TextRun({ text: "\t", font: FONT, size: BODY })],
      tabStops: [{ type: TabStopType.LEFT, position: 1440 }, { type: TabStopType.LEFT, position: 6480, leader: "underscore" }] }),
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Title:\t", font: FONT, size: BODY }), new TextRun({ text: "\t", font: FONT, size: BODY })],
      tabStops: [{ type: TabStopType.LEFT, position: 1440 }, { type: TabStopType.LEFT, position: 6480, leader: "underscore" }] }),
    new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "Date:\t", font: FONT, size: BODY }), new TextRun({ text: "\t", font: FONT, size: BODY })],
      tabStops: [{ type: TabStopType.LEFT, position: 1440 }, { type: TabStopType.LEFT, position: 6480, leader: "underscore" }] }),
  ];
}

// ---- content ---------------------------------------------------------------
const children = [];

// Cover / caption
children.push(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 240 },
    children: [new TextRun({ text: "DRAFT — FOR ATTORNEY REVIEW — NOT FOR EXECUTION", font: FONT, size: 20, bold: true, color: "C00000" })] }),
  centered("JOINT VENTURE OPERATING AGREEMENT", 32, true, 80),
  centered("of", BODY, false, 80),
  centered("TEPER–KGK CARD VENTURE", 28, true, 80),
  centered("(a joint venture between Teper Holdings and Group KGK Inc.)", BODY, false, 240),
  centered("Dated as of [__________], 2026", BODY, false, 360),
);

children.push(
  para("THIS JOINT VENTURE OPERATING AGREEMENT (this “Agreement”) is entered into as of [__________], 2026 (the “Effective Date”), by and between:"),
  sub("1", "TEPER HOLDINGS, a [State] [limited liability company / corporation] having its principal office at [Address] (“Teper Holdings”); and"),
  sub("2", "GROUP KGK INC., a [State] corporation having its principal office at [Address] (“KGK”)."),
  para("Teper Holdings and KGK are each referred to herein as a “Member” and collectively as the “Members.”"),
);

// Recitals
children.push(centered("RECITALS", BODY, true, 160));
children.push(
  para("A.\tThe Members wish to jointly acquire, hold, manage, and ultimately sell one or more collectible trading cards, including the card(s) described on Schedule A (together with any additional cards acquired pursuant to this Agreement, the “Cards”), for investment purposes and for the shared profit of the Members."),
  para("B.\tTeper Holdings has funded, or will fund, one hundred percent (100%) of the purchase price and initial acquisition costs of the Cards on behalf of the venture (the “Initial Contribution”), and KGK has agreed to contribute its share of the capital at a later date on the terms set forth in Article III (the “Deferred Contribution”)."),
  para("C.\tThe Members wish to set forth in writing their respective rights, obligations, ownership interests, and the manner in which the Cards will be held, managed, and sold, and how proceeds will be distributed."),
  para("NOW, THEREFORE, in consideration of the mutual covenants set forth herein and other good and valuable consideration, the receipt and sufficiency of which are hereby acknowledged, the Members agree as follows:"),
);

// Article I
children.push(article("I", "Formation, Name, Purpose, and Term"));
children.push(
  sec("1.1", "Formation", "The Members hereby form a joint venture (the “Venture”) under the laws of the State of [State] for the limited purpose described in Section 1.3. [The Venture is an unincorporated joint venture that the Members intend to be treated as a general partnership solely for U.S. federal and applicable state income tax purposes.] [ALTERNATIVE: The Members shall cause a limited liability company to be formed under the name set forth in Section 1.2 in the State of [State], and this Agreement shall constitute the operating agreement of such limited liability company.]"),
  note("Attorney to confirm the entity choice. Holding the Cards through an unincorporated joint venture is simplest but exposes each Member to joint and several liability for Venture obligations. Forming a single-purpose LLC provides a liability shield and a clean owner of record for the Cards, at the cost of a filing fee and annual compliance. If an LLC is formed, the bracketed alternative language should be used and the Members become “members” of that LLC."),
  sec("1.2", "Name", "The business of the Venture shall be conducted under the name “Teper–KGK Card Venture” or such other name as the Members may unanimously select."),
  sec("1.3", "Purpose", "The sole purpose of the Venture is to (a) acquire the Cards, (b) hold, store, insure, authenticate, grade, and otherwise preserve and protect the Cards, (c) market and sell the Cards at such time and on such terms as are determined in accordance with this Agreement, and (d) engage in any activities necessary or incidental to the foregoing. The Venture shall not engage in any other business without the unanimous written consent of the Members."),
  sec("1.4", "Term", "The Venture shall commence on the Effective Date and shall continue until the earlier of (a) the sale or other disposition of all of the Cards and the distribution of all net proceeds in accordance with Article VII, or (b) dissolution pursuant to Article XI."),
  sec("1.5", "Principal Office", "The principal office of the Venture shall be located at [Address], or such other place as the Members may agree."),
  sec("1.6", "No Partnership Beyond the Venture", "Except as expressly provided herein with respect to the Cards, nothing in this Agreement shall be construed to create a partnership, agency, or fiduciary relationship between the Members with respect to any other business, asset, or activity of either Member. Neither Member shall have authority to bind the other Member, or the Venture, except as expressly provided in this Agreement."),
);

// Article II
children.push(article("II", "Definitions"));
children.push(
  para("Capitalized terms used in this Agreement have the meanings set forth below or elsewhere in this Agreement:"),
  sub("a", "“Acquisition Cost” means the aggregate purchase price of the Cards together with all buyer’s premiums, platform or marketplace fees, sales taxes, shipping, insurance in transit, and other direct costs of acquiring the Cards, in each case as itemized on Schedule B."),
  sub("b", "“Capital Account” means the separate account maintained for each Member pursuant to Section 3.6."),
  sub("c", "“Capital Contribution” means, as to any Member, the total amount of cash and the agreed fair market value of property actually contributed to the Venture by such Member, including the Initial Contribution and the Deferred Contribution when and to the extent actually paid."),
  sub("d", "“Cards” has the meaning set forth in Recital A."),
  sub("e", "“Deferred Contribution” has the meaning set forth in Section 3.2."),
  sub("f", "“Deferred Contribution Deadline” means [__________], 2026, or such later date as Teper Holdings may agree in writing in its sole discretion."),
  sub("g", "“Initial Contribution” has the meaning set forth in Section 3.1."),
  sub("h", "“Major Decision” has the meaning set forth in Section 5.3."),
  sub("i", "“Managing Member” means Teper Holdings, or any successor Managing Member appointed pursuant to Section 5.2."),
  sub("j", "“Net Proceeds” means the gross proceeds received by the Venture from any sale, consignment, insurance recovery, or other disposition of any Card, less all seller’s fees, commissions, platform fees, shipping, taxes, and other costs of such disposition."),
  sub("k", "“Percentage Interest” means, as to each Member, the percentage set forth opposite such Member’s name on Schedule B, as adjusted from time to time pursuant to this Agreement."),
  sub("l", "“Target Percentage Interest” means the Percentage Interest each Member is intended to hold upon full funding of both the Initial Contribution and the Deferred Contribution, being [fifty percent (50%)] for Teper Holdings and [fifty percent (50%)] for KGK."),
  sub("m", "“Venture Expenses” has the meaning set forth in Section 6.1."),
);

// Article III — capital
children.push(article("III", "Capital Contributions"));
children.push(
  sec("3.1", "Initial Contribution by Teper Holdings", "Teper Holdings has contributed, or shall contribute, to the Venture cash in the amount of [$__________], representing one hundred percent (100%) of the Acquisition Cost of the Cards identified on Schedule A as of the Effective Date (the “Initial Contribution”). The Initial Contribution is itemized on Schedule B. KGK acknowledges that the Cards were purchased with funds provided solely by Teper Holdings and that KGK has made no Capital Contribution as of the Effective Date."),
  sec("3.2", "Deferred Contribution by KGK", "KGK shall contribute to the Venture, on or before the Deferred Contribution Deadline, cash in an amount equal to [fifty percent (50%)] of the Acquisition Cost, being [$__________] (the “Deferred Contribution”), [together with interest on such amount at the rate of [___]% per annum, simple interest, from the date of the Initial Contribution to the date the Deferred Contribution is paid in full]. The Deferred Contribution shall be paid to Teper Holdings by wire transfer, ACH, or such other method as Teper Holdings may reasonably direct, and shall be treated as (a) a Capital Contribution by KGK to the Venture and (b) a simultaneous return of capital to Teper Holdings, such that upon payment in full each Member’s Capital Contribution equals its Target Percentage Interest of the Acquisition Cost. [KGK may pay the Deferred Contribution in installments in accordance with the schedule set forth on Schedule C.]"),
  note("Confirm the amount and the deadline. The Members should decide whether Teper Holdings’ carry of KGK’s share bears interest. If no interest is intended, delete the bracketed interest language. An installment schedule is optional and can be attached as Schedule C."),
  sec("3.3", "Percentage Interests Pending Deferred Contribution", "From the Effective Date until the Deferred Contribution is paid in full, the Members’ Percentage Interests shall be as follows: Teper Holdings, [one hundred percent (100%)]; KGK, [zero percent (0%)]; provided, however, that KGK shall hold a contractual right (the “Contribution Right”) to acquire its Target Percentage Interest by timely paying the Deferred Contribution in full in accordance with Section 3.2. Upon receipt by Teper Holdings of the Deferred Contribution in full, the Percentage Interests shall automatically adjust to the Target Percentage Interests, effective as of the date of such payment, and the Managing Member shall update Schedule B accordingly."),
  note("ALTERNATIVE (Members’ choice): The Members may instead prefer that KGK hold its Target Percentage Interest from the Effective Date, with the unpaid Deferred Contribution treated as a debt owed by KGK to the Venture that is deducted from KGK’s distributions under Section 7.2. That approach is more favorable to KGK; the approach drafted above is more protective of Teper Holdings, which has funded the entire purchase. Attorney should select one approach with the clients."),
  sec("3.4", "Failure to Pay Deferred Contribution", "If KGK fails to pay the Deferred Contribution in full on or before the Deferred Contribution Deadline, then, at the election of Teper Holdings by written notice to KGK, any one or more of the following shall apply:"),
  sub("a", "Partial Funding. If KGK has paid a portion of the Deferred Contribution, KGK’s Percentage Interest shall equal the percentage that KGK’s actual Capital Contribution bears to the total Capital Contributions of all Members, and the Contribution Right as to the unpaid balance shall terminate;"),
  sub("b", "Termination of Contribution Right. The Contribution Right shall terminate in its entirety, KGK shall have no Percentage Interest in the Venture or the Cards, and any partial payment made by KGK shall be refunded to KGK, without interest, within thirty (30) days; or"),
  sub("c", "Extension. Teper Holdings may extend the Deferred Contribution Deadline for such period and on such terms (including the accrual of interest) as Teper Holdings may determine in its sole discretion."),
  para("The remedies set forth in this Section 3.4 are in addition to, and not in lieu of, any other remedies available to Teper Holdings at law or in equity."),
  sec("3.5", "Additional Capital Contributions", "No Member shall be required to make any Capital Contribution beyond those described in Sections 3.1 and 3.2. If the Members unanimously determine that additional capital is required (for example, to acquire additional Cards, or to pay grading, insurance, or storage costs in excess of Venture cash), the Members shall contribute such additional capital in proportion to their Percentage Interests, unless otherwise agreed in writing. If one Member advances funds on behalf of a non-contributing Member, such advance shall be treated as a Member Loan under Section 6.3."),
  sec("3.6", "Capital Accounts", "A separate Capital Account shall be maintained for each Member in accordance with Treasury Regulations Section 1.704-1(b)(2)(iv). Each Member’s Capital Account shall be (a) increased by its Capital Contributions and its allocable share of Venture income and gain, and (b) decreased by distributions to it and its allocable share of Venture losses and deductions."),
  sec("3.7", "No Interest; No Withdrawal", "Except as expressly provided in Section 3.2, no Member shall be entitled to interest on its Capital Contribution or Capital Account, and no Member may withdraw any part of its Capital Contribution or receive any distribution except as provided in this Agreement."),
);

// Article IV — ownership of the cards
children.push(article("IV", "Ownership, Title, Custody, and Insurance of the Cards"));
children.push(
  sec("4.1", "Ownership", "The Cards are, and shall remain, property of the Venture. Each Member’s interest in the Cards is an undivided beneficial interest equal to its Percentage Interest as in effect from time to time. No Member shall have any right to partition the Cards or to require distribution of any Card in kind."),
  sec("4.2", "Title and Registration", "Legal title to, and possession of, the Cards shall be held by [Teper Holdings, as nominee and custodian for the Venture / the LLC formed pursuant to Section 1.1]. Any marketplace, auction, vault, or grading-company account through which a Card is purchased, registered, stored, or sold (including any Whatnot, PSA, or vault account) shall be identified on Schedule A, and the Managing Member shall maintain such accounts in good standing. The Managing Member shall not permit any Card to be titled or registered in the name of any person other than the Venture or the nominee described above."),
  sec("4.3", "Custody and Storage", "The Cards shall be stored [at [Address] / in a third-party vault or safe-deposit facility approved by the Members] in a manner reasonably designed to preserve their condition and grade, including, for graded Cards, maintaining the original grading-company holder intact and unopened. The Managing Member shall provide the other Member with reasonable access to inspect the Cards upon not less than [five (5)] business days’ notice."),
  sec("4.4", "Insurance", "The Managing Member shall [use commercially reasonable efforts to] procure and maintain, as a Venture Expense, collectibles insurance covering the Cards at not less than their aggregate Acquisition Cost [or current appraised value, if greater], naming the Venture (or its nominee) as insured. Any insurance proceeds received with respect to a Card shall constitute Net Proceeds and be distributed in accordance with Article VII."),
  sec("4.5", "Grading, Authentication, and Alteration", "No Card shall be removed from its grading holder, submitted for regrading, crossover, review, or reholdering, cleaned, altered, or otherwise physically modified without the unanimous written consent of the Members."),
  sec("4.6", "No Encumbrance", "No Member shall pledge, mortgage, hypothecate, or otherwise encumber any Card or any interest therein, or permit any lien to attach to any Card, without the unanimous written consent of the Members."),
  sec("4.7", "Risk of Loss", "Any loss, theft, damage, or destruction of a Card that is not covered by insurance shall be borne by the Members in proportion to their Percentage Interests, except to the extent caused by the gross negligence, willful misconduct, or breach of this Agreement by a Member, in which case such Member shall bear such loss."),
);

// Article V — management
children.push(article("V", "Management"));
children.push(
  sec("5.1", "Management by the Members", "Except as delegated to the Managing Member under Section 5.2, the business and affairs of the Venture shall be managed by the Members jointly. Each Member shall have voting power equal to its Percentage Interest; provided that, notwithstanding the Percentage Interests in effect at any time, all Major Decisions shall require the unanimous written consent of both Members [for so long as KGK’s Contribution Right remains in effect or KGK holds a Percentage Interest]."),
  sec("5.2", "Managing Member", "Teper Holdings is hereby designated as the Managing Member. The Managing Member shall be responsible for the day-to-day administration of the Venture, including custody and storage of the Cards, maintenance of marketplace and grading accounts, payment of Venture Expenses, maintenance of the books and records, and preparation of tax filings. The Managing Member shall act in good faith and in the best interests of the Venture, and shall keep the other Member reasonably informed of all material matters. The Managing Member may be removed and replaced only by unanimous written consent of the Members or upon the Managing Member’s material breach of this Agreement that remains uncured thirty (30) days after written notice."),
  sec("5.3", "Major Decisions", "None of the following actions (each, a “Major Decision”) shall be taken by the Venture or any Member on its behalf without the unanimous written consent of the Members:"),
  sub("a", "the sale, consignment, trade, auction listing, or other disposition of any Card, including the selection of the sales channel (e.g., auction house, marketplace, private sale) and the establishment or modification of any reserve, minimum, or “buy it now” price;"),
  sub("b", "the acquisition of any additional card or other asset;"),
  sub("c", "any grading, regrading, crossover, or physical alteration of a Card as described in Section 4.5;"),
  sub("d", "any encumbrance of a Card or any borrowing by the Venture;"),
  sub("e", "any change in the storage location or custodian of the Cards;"),
  sub("f", "the admission of any new Member or the transfer of any interest in the Venture;"),
  sub("g", "any Venture Expense or series of related Venture Expenses in excess of [$________] not already contemplated by an approved budget;"),
  sub("h", "any amendment of this Agreement or of any Schedule hereto (other than ministerial updates to Schedules A and B by the Managing Member to reflect transactions approved hereunder);"),
  sub("i", "the dissolution of the Venture other than as provided in Article XI; and"),
  sub("j", "any lawsuit, settlement, or insurance claim on behalf of the Venture."),
  sec("5.4", "Written Consent; Electronic Communications", "Any consent or approval of the Members required under this Agreement may be given in writing, including by email or text message to the addresses set forth in Section 13.2, provided the communication clearly identifies the action approved. The Managing Member shall retain copies of all such consents with the Venture’s records."),
  sec("5.5", "Deadlock", "If the Members are unable to agree on a Major Decision within [thirty (30)] days after either Member has proposed such decision in writing (a “Deadlock”), either Member may invoke the buy-sell procedure set forth in Section 8.3."),
  sec("5.6", "Time Commitment; Other Activities", "No Member shall be required to devote any particular amount of time to the Venture. Each Member may engage in other business ventures of any nature, including the purchase, sale, and collection of trading cards for its own account, and neither the Venture nor the other Member shall have any right in or to such other ventures or the income derived therefrom; provided that no Member shall use Venture funds or Cards for any purpose other than the Venture’s business."),
  sec("5.7", "Compensation", "No Member shall receive any salary, fee, or other compensation for services rendered to the Venture, except that the Managing Member shall be reimbursed for Venture Expenses as provided in Article VI [and shall be entitled to a management fee of [___]% of Net Proceeds upon the sale of each Card]."),
);

// Article VI — expenses
children.push(article("VI", "Venture Expenses and Member Loans"));
children.push(
  sec("6.1", "Venture Expenses", "“Venture Expenses” means all reasonable out-of-pocket costs incurred in connection with the Venture’s purpose, including Acquisition Costs, grading and authentication fees, storage and vault fees, insurance premiums, shipping and handling, marketplace listing and seller fees, appraisal fees, sales and use taxes, and reasonable accounting, tax preparation, and legal fees of the Venture. Venture Expenses shall be borne by the Members in proportion to their Percentage Interests."),
  sec("6.2", "Payment and Reimbursement", "The Managing Member may pay Venture Expenses from Venture cash or, if Venture cash is insufficient, may advance such amounts on behalf of the Venture. All advances by a Member shall be documented on the expense ledger attached as Schedule C and shall be reimbursed to such Member, with priority over any other distribution, in accordance with Section 7.2(a). Any Venture Expense in excess of the threshold in Section 5.3(g) shall require the prior written consent of both Members."),
  sec("6.3", "Member Loans", "Any amount advanced by a Member on behalf of the Venture that exceeds such Member’s proportionate share of Venture Expenses (a “Member Loan”) shall bear interest at [___]% per annum, simple interest, from the date advanced until repaid, and shall be repaid in accordance with Section 7.2(a) before any distribution of capital or profits. A Member Loan shall not be treated as a Capital Contribution and shall not affect the Percentage Interests."),
  sec("6.4", "Bank Account", "The Managing Member may, and upon the request of either Member shall, open a segregated bank or payment account in the name of the Venture (or the Managing Member as nominee for the Venture) through which all Venture receipts and disbursements shall be made. Venture funds shall not be commingled with the personal or business funds of any Member."),
);

// Article VII — distributions & allocations
children.push(article("VII", "Distributions and Allocations"));
children.push(
  sec("7.1", "Distributions of Net Proceeds", "Net Proceeds from the sale or other disposition of any Card, and any other cash available for distribution as determined by the Managing Member, shall be distributed to the Members within [fifteen (15)] business days after receipt by the Venture of cleared funds, in the order of priority set forth in Section 7.2."),
  sec("7.2", "Order of Priority", "Distributions shall be made in the following order:"),
  sub("a", "First, to the Members, pro rata in proportion to and to the extent of any outstanding Member Loans and unreimbursed Venture Expense advances (including accrued interest), until all such amounts have been paid in full;"),
  sub("b", "Second, if the Deferred Contribution has not been paid in full at the time of distribution, to Teper Holdings until it has received an amount equal to the unpaid balance of the Deferred Contribution [together with any accrued interest thereon], which amount shall be credited against, and reduce, KGK’s Deferred Contribution obligation and shall be treated for purposes of Section 3.3 as a payment of the Deferred Contribution by KGK;"),
  sub("c", "Third, to the Members in proportion to, and to the extent of, their respective unreturned Capital Contributions; and"),
  sub("d", "Thereafter, to the Members in proportion to their Percentage Interests as in effect on the date of distribution."),
  note("Section 7.2(b) is only relevant under the “ALTERNATIVE” approach described in the Drafting Note to Section 3.3, or if a Card is sold before the Deferred Contribution Deadline. Under the primary approach in Section 3.3, KGK has no Percentage Interest until it pays, so a sale before payment would be distributed entirely to Teper Holdings under clauses (c) and (d). The Members should confirm which outcome they intend if a Card is sold before KGK funds."),
  sec("7.3", "Distributions in Kind", "No Member shall have the right to demand or receive any Card or other property in kind. The Members may, by unanimous written consent, agree that a Card be distributed in kind to one Member at an agreed value, with an appropriate cash adjustment to the other Member."),
  sec("7.4", "Tax Allocations", "Items of Venture income, gain, loss, deduction, and credit shall be allocated among the Members in a manner consistent with the distributions provided for in Section 7.2 and in compliance with Section 704(b) of the Internal Revenue Code and the Treasury Regulations thereunder. Gain or loss on the sale of a Card shall be allocated in proportion to the Members’ Percentage Interests in effect on the date of sale, subject to adjustment so that Capital Accounts, after all allocations, equal the amounts distributable to the Members under Section 7.2."),
  sec("7.5", "Limitations", "No distribution shall be made if, after giving effect thereto, the Venture would be unable to pay its debts as they become due or the distribution would otherwise violate applicable law."),
);

// Article VIII — sale process, ROFR, buy-sell
children.push(article("VIII", "Sale of the Cards; Buy-Sell"));
children.push(
  sec("8.1", "Sale Strategy", "The Members acknowledge that the Cards are held for investment with a target holding period of [__________], subject to market conditions. Either Member may at any time propose in writing the sale of any Card, specifying the proposed sales channel, timing, and minimum price. Such proposal shall constitute a Major Decision under Section 5.3(a)."),
  sec("8.2", "Right of First Offer", "If a sale of a Card has been approved as a Major Decision, or if either Member wishes to exit the Venture with respect to a Card, the other Member (the “Offeree Member”) shall have the right, exercisable by written notice within [ten (10)] business days after receipt of the sale proposal, to purchase the proposing Member’s entire interest in such Card for cash at a price equal to the proposing Member’s Percentage Interest multiplied by the greater of (a) the minimum price stated in the sale proposal and (b) the most recent [PSA / third-party] appraised value or comparable-sales value reasonably agreed by the Members. If the Offeree Member does not exercise such right, the Card may be sold to a third party in accordance with the approved proposal."),
  sec("8.3", "Buy-Sell upon Deadlock", "If a Deadlock occurs under Section 5.5, either Member (the “Initiating Member”) may deliver to the other Member (the “Responding Member”) a written notice stating a single cash value for one hundred percent (100%) of the Cards (the “Stated Value”). The Responding Member shall, within [twenty (20)] business days after receipt of such notice, elect in writing either (a) to purchase the Initiating Member’s entire interest in the Venture for an amount equal to the Initiating Member’s Percentage Interest multiplied by the Stated Value, or (b) to sell the Responding Member’s entire interest in the Venture to the Initiating Member for an amount equal to the Responding Member’s Percentage Interest multiplied by the Stated Value. Failure to elect within such period shall be deemed an election under clause (b). The closing of any such purchase shall occur within [thirty (30)] days after such election, against payment in immediately available funds, and any outstanding Member Loans and unpaid Deferred Contribution amounts shall be settled at such closing."),
  note("If KGK has not yet paid the Deferred Contribution when a buy-sell is triggered, its “Percentage Interest” under the primary approach is 0%, which would make the buy-sell meaningless for KGK. Attorney should consider whether the buy-sell should be available only after the Deferred Contribution is paid, or whether KGK’s interest for buy-sell purposes should be measured net of the unpaid Deferred Contribution."),
  sec("8.4", "Cooperation", "Each Member shall execute such bills of sale, transfer forms, marketplace listings, and other documents as may be reasonably necessary to consummate any sale approved in accordance with this Agreement."),
);

// Article IX — books, records, tax
children.push(article("IX", "Books, Records, Reporting, and Tax Matters"));
children.push(
  sec("9.1", "Books and Records", "The Managing Member shall maintain complete and accurate books and records of the Venture, including (a) this Agreement and all amendments, (b) Schedules A, B, and C as updated from time to time, (c) purchase receipts, invoices, and grading certifications for each Card, (d) the expense ledger, (e) all written consents of the Members, and (f) bank and marketplace account statements. Such books and records shall be available for inspection and copying by either Member upon reasonable notice."),
  sec("9.2", "Reporting", "The Managing Member shall provide the other Member with (a) written notice of any purchase, sale, insurance claim, or loss involving any Card within [five (5)] business days after the occurrence thereof, and (b) an annual statement of Venture Expenses, distributions, and Capital Accounts within [sixty (60)] days after the end of each calendar year."),
  sec("9.3", "Tax Classification", "The Members intend that the Venture be classified as a partnership for U.S. federal and applicable state income tax purposes, and no Member shall make any election to the contrary. The Members acknowledge that gain on the sale of collectibles held for more than one year may be subject to the collectibles capital gains rate under Section 1(h)(5) of the Internal Revenue Code, and each Member shall be responsible for its own tax liabilities arising from the Venture."),
  note("Attorney and/or CPA should confirm whether a partnership return (Form 1065) and Schedule K-1s will be required, or whether the arrangement may qualify for an election out of subchapter K under IRC § 761(a) as a mere co-ownership of investment property. This affects reporting cost and complexity."),
  sec("9.4", "Partnership Representative", "Teper Holdings is designated as the “partnership representative” of the Venture within the meaning of Section 6223 of the Internal Revenue Code, and shall keep the other Member reasonably informed of any tax audit or proceeding. Any settlement or election that would disproportionately affect KGK shall require KGK’s prior written consent, not to be unreasonably withheld."),
  sec("9.5", "Fiscal Year; Accounting Method", "The fiscal year of the Venture shall be the calendar year. The books of the Venture shall be kept on the cash method of accounting, unless otherwise required by law."),
);

// Article X — transfers
children.push(article("X", "Transfer of Interests"));
children.push(
  sec("10.1", "Restrictions on Transfer", "No Member shall sell, assign, pledge, or otherwise transfer all or any part of its interest in the Venture or the Cards (a “Transfer”), whether voluntarily, involuntarily, or by operation of law, without the prior written consent of the other Member, which may be withheld in such Member’s sole discretion. Any purported Transfer in violation of this Article X shall be null and void."),
  sec("10.2", "Permitted Transfers", "Notwithstanding Section 10.1, a Member may Transfer its interest to an entity wholly owned and controlled by such Member or by such Member’s principal, provided that (a) the transferee agrees in writing to be bound by this Agreement, and (b) the transferring Member remains liable for its obligations hereunder."),
  sec("10.3", "Change of Control", "Any change in the ownership or control of a Member that is an entity shall be deemed a Transfer of such Member’s interest for purposes of this Article X."),
);

// Article XI — dissolution
children.push(article("XI", "Dissolution and Winding Up"));
children.push(
  sec("11.1", "Events of Dissolution", "The Venture shall be dissolved upon the first to occur of: (a) the sale or other disposition of all of the Cards and the distribution of all Net Proceeds; (b) the unanimous written agreement of the Members; (c) the purchase by one Member of the other Member’s entire interest pursuant to Section 8.2 or 8.3; or (d) any event that makes it unlawful for the business of the Venture to be carried on."),
  sec("11.2", "Winding Up", "Upon dissolution, the Managing Member (or, if the Managing Member is in breach, the other Member) shall wind up the affairs of the Venture, sell or otherwise dispose of the Cards on commercially reasonable terms (or distribute them in kind by unanimous consent), pay or provide for all Venture liabilities, and distribute the remaining assets in the order set forth in Section 7.2."),
  sec("11.3", "Bankruptcy or Dissolution of a Member", "The bankruptcy, insolvency, or dissolution of a Member shall not dissolve the Venture, but the other Member shall have the right, exercisable within [sixty (60)] days after receiving notice of such event, to purchase the affected Member’s interest at its fair market value as determined by an independent appraiser mutually selected by the Members (or, failing agreement, by the [American Arbitration Association])."),
);

// Article XII — reps, indemnification
children.push(article("XII", "Representations, Warranties, and Indemnification"));
children.push(
  sec("12.1", "Mutual Representations", "Each Member represents and warrants to the other that: (a) it is duly organized, validly existing, and in good standing under the laws of its jurisdiction of organization; (b) it has full power and authority to enter into and perform this Agreement, and this Agreement has been duly authorized, executed, and delivered by it; (c) the execution and performance of this Agreement do not conflict with any agreement, order, or law binding on it; and (d) it is entering into this Agreement for investment purposes and not with a view to resale or distribution of its interest in violation of applicable securities laws."),
  sec("12.2", "Representations of Teper Holdings", "Teper Holdings represents and warrants that (a) it has paid the Initial Contribution in the amount set forth in Section 3.1 and has furnished KGK with true and complete copies of the purchase receipts for the Cards, and (b) to its knowledge, the Cards are authentic, the grading-company holders are intact, and Teper Holdings has good title to the Cards free of any lien or claim."),
  sec("12.3", "Representations of KGK", "KGK represents and warrants that it has, and will have as of the Deferred Contribution Deadline, sufficient funds available to pay the Deferred Contribution in full, and that its principal, [Name], has authority to bind KGK."),
  sec("12.4", "Limitation of Liability", "No Member shall be liable to the Venture or the other Member for any act or omission performed or omitted in good faith on behalf of the Venture and in a manner reasonably believed to be within the scope of authority conferred by this Agreement, except for acts or omissions constituting gross negligence, willful misconduct, fraud, or a material breach of this Agreement."),
  sec("12.5", "Indemnification", "Each Member (the “Indemnifying Member”) shall indemnify, defend, and hold harmless the other Member and the Venture from and against any loss, liability, damage, or expense (including reasonable attorneys’ fees) arising out of the Indemnifying Member’s gross negligence, willful misconduct, fraud, or material breach of this Agreement."),
  sec("12.6", "No Investment Advice; Risk Acknowledgment", "Each Member acknowledges that (a) the value of collectible trading cards is volatile and speculative, (b) no Member has made any representation regarding the future value or liquidity of the Cards, (c) it has made its own independent evaluation of the merits and risks of the Venture, and (d) it may lose all or a portion of its Capital Contribution."),
);

// Article XIII — misc
children.push(article("XIII", "General Provisions"));
children.push(
  sec("13.1", "Governing Law", "This Agreement shall be governed by and construed in accordance with the laws of the State of [State], without regard to its conflict-of-laws principles."),
  sec("13.2", "Notices", "All notices under this Agreement shall be in writing and shall be deemed given when delivered personally, sent by nationally recognized overnight courier, or sent by email with confirmation of transmission, to the addresses set forth below or such other address as a Member may designate by notice: Teper Holdings: [Address] / [Email], Attn: [Matthew Teper]; KGK: [Address] / [Email], Attn: [Name]."),
  sec("13.3", "Dispute Resolution", "The Members shall attempt in good faith to resolve any dispute arising out of this Agreement through direct negotiation for a period of [thirty (30)] days. Any dispute not so resolved shall be [submitted to binding arbitration administered by the [American Arbitration Association] under its Commercial Arbitration Rules, with the seat of arbitration in [City, State] / brought exclusively in the state or federal courts located in [County, State], and each Member consents to the personal jurisdiction of such courts]. The prevailing party in any such proceeding shall be entitled to recover its reasonable attorneys’ fees and costs."),
  sec("13.4", "Entire Agreement; Amendment", "This Agreement, including its Schedules, constitutes the entire agreement of the Members with respect to the subject matter hereof and supersedes all prior agreements and understandings, oral or written. This Agreement may be amended only by a written instrument signed by both Members."),
  sec("13.5", "Waiver", "No waiver of any provision of this Agreement shall be effective unless in writing and signed by the Member against whom the waiver is asserted, and no waiver shall constitute a waiver of any other or subsequent breach."),
  sec("13.6", "Severability", "If any provision of this Agreement is held invalid or unenforceable, the remaining provisions shall continue in full force and effect, and the invalid provision shall be reformed to the minimum extent necessary to make it enforceable."),
  sec("13.7", "Binding Effect; No Third-Party Beneficiaries", "This Agreement shall be binding upon and inure to the benefit of the Members and their respective permitted successors and assigns. Nothing in this Agreement shall confer any rights on any person other than the Members."),
  sec("13.8", "Further Assurances", "Each Member shall execute and deliver such further documents and take such further actions as may be reasonably necessary to carry out the purposes of this Agreement."),
  sec("13.9", "Counterparts; Electronic Signatures", "This Agreement may be executed in counterparts, each of which shall be deemed an original, and all of which together shall constitute one instrument. Signatures delivered by PDF, DocuSign, or other electronic means shall be effective as originals."),
  sec("13.10", "Independent Counsel", "Each Member acknowledges that it has been advised to seek, and has had the opportunity to obtain, independent legal and tax advice regarding this Agreement, and that [counsel for Teper Holdings prepared this Agreement and does not represent KGK]."),
  sec("13.11", "Headings; Interpretation", "Headings are for convenience only and shall not affect interpretation. The words “include” and “including” are deemed to be followed by “without limitation.” This Agreement shall be construed as if drafted jointly by the Members."),
);

// Signatures
children.push(
  blank(200),
  para("IN WITNESS WHEREOF, the Members have executed this Joint Venture Operating Agreement as of the Effective Date."),
  ...sigBlock("TEPER HOLDINGS", "a [State] [limited liability company / corporation]"),
  ...sigBlock("GROUP KGK INC.", "a [State] corporation"),
);

// Schedule A
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(centered("SCHEDULE A", 24, true, 60), centered("THE CARDS", BODY, true, 200));
children.push(para("The following Cards are subject to this Agreement as of the Effective Date. The Managing Member shall update this Schedule to reflect any additional Cards acquired, or any Card sold, in accordance with the Agreement."));
children.push(table(
  ["#", "Description", "Grade / Cert. No.", "Seller / Platform", "Purchase Date", "Purchase Price"],
  [
    ["1", "[2025 Topps Chrome Update — LeBron James — Moment in Time — White Geo Refractor]", "[PSA 10 GEM MT / Cert. No. 1756343__ — CONFIRM]", "[debutsports_ via Whatnot livestream — CONFIRM]", "[__/__/2026]", "[$________]"],
    ["2", "[2025 Topps Chrome Update — Stephen Curry — Moment in Time — Red Geo Refractor, #[2]/5]", "[PSA 10 GEM MT / Cert. No. 175141683 — CONFIRM]", "[debutsports_ via Whatnot livestream — CONFIRM]", "[__/__/2026]", "[$________]"],
    ["3", "[__________]", "[__________]", "[__________]", "[__/__/____]", "[$________]"],
  ],
  [400, 3200, 1900, 1800, 1100, 960],
));
children.push(blank(160));
children.push(note("Card descriptions above were transcribed from the purchase screenshot supplied by Teper Holdings and must be verified against the PSA certification labels and the Whatnot order confirmation before execution. Confirm the LeBron James certification number (partially obscured), the Curry serial number, and whether both cards (or only one) were acquired for the Venture."));
children.push(para("Accounts associated with the Cards:", { run: { bold: true }, para: { alignment: AlignmentType.LEFT } }));
children.push(
  bullet("Marketplace / auction account: [Whatnot — account holder: __________]"),
  bullet("Grading company account: [PSA — account holder: __________]"),
  bullet("Vault / storage: [__________]"),
  bullet("Insurance policy: [Carrier / Policy No. __________]"),
);

// Schedule B
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(centered("SCHEDULE B", 24, true, 60), centered("CAPITAL CONTRIBUTIONS AND PERCENTAGE INTERESTS", BODY, true, 200));
children.push(para("Part 1 — Acquisition Cost (funded 100% by Teper Holdings as the Initial Contribution):", { run: { bold: true }, para: { alignment: AlignmentType.LEFT } }));
children.push(table(
  ["Item", "Amount"],
  [
    ["Purchase price — Card 1", "[$________]"],
    ["Purchase price — Card 2", "[$________]"],
    ["Buyer’s premium / platform fees", "[$________]"],
    ["Sales tax", "[$________]"],
    ["Shipping and transit insurance", "[$________]"],
    ["Other acquisition costs (describe): [__________]", "[$________]"],
    ["TOTAL ACQUISITION COST / INITIAL CONTRIBUTION", "[$________]"],
  ],
  [6800, 2560],
));
children.push(blank(200));
children.push(para("Part 2 — Capital Contributions and Percentage Interests:", { run: { bold: true }, para: { alignment: AlignmentType.LEFT } }));
children.push(table(
  ["Member", "Capital Contribution as of Effective Date", "Percentage Interest as of Effective Date", "Deferred Contribution Owed", "Target Percentage Interest (upon payment in full)"],
  [
    ["Teper Holdings", "[$________] (100% of Acquisition Cost)", "[100%]", "—", "[50%]"],
    ["Group KGK Inc.", "$0.00", "[0%] (Contribution Right only)", "[$________] (50% of Acquisition Cost) [+ interest at ___%]", "[50%]"],
    ["Total", "[$________]", "100%", "", "100%"],
  ],
  [1700, 2100, 1900, 1900, 1760],
));
children.push(blank(160));
children.push(para("Deferred Contribution Deadline: [__________], 2026.", { para: { alignment: AlignmentType.LEFT } }));
children.push(para("Schedule B last updated: [__________] by the Managing Member.", { para: { alignment: AlignmentType.LEFT } }));

// Schedule C
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(centered("SCHEDULE C", 24, true, 60), centered("DEFERRED CONTRIBUTION PAYMENT SCHEDULE AND EXPENSE LEDGER", BODY, true, 200));
children.push(para("Part 1 — Deferred Contribution Payment Schedule (optional; delete if paid in a single installment):", { run: { bold: true }, para: { alignment: AlignmentType.LEFT } }));
children.push(table(
  ["Installment", "Due Date", "Amount", "Date Paid", "Confirmed By (Teper Holdings)"],
  [["1", "[__/__/2026]", "[$________]", "", ""], ["2", "[__/__/2026]", "[$________]", "", ""], ["3", "[__/__/2026]", "[$________]", "", ""], ["Total", "", "[$________]", "", ""]],
  [1300, 1700, 1700, 1700, 2960],
));
children.push(blank(200));
children.push(para("Part 2 — Venture Expense and Member Advance Ledger:", { run: { bold: true }, para: { alignment: AlignmentType.LEFT } }));
children.push(table(
  ["Date", "Description", "Paid By", "Amount", "Reimbursed (Y/N / Date)"],
  [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]],
  [1300, 3500, 1500, 1300, 1760],
));

// Attorney review memo
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(centered("ATTORNEY REVIEW MEMORANDUM", 24, true, 60), centered("(Not part of the Agreement — remove before execution)", BODY, false, 200));
children.push(para("This draft was prepared from the following facts supplied by Teper Holdings and should be verified: (1) Teper Holdings and Group KGK Inc. are jointly purchasing graded trading cards as an investment; (2) Teper Holdings funded 100% of the purchase; (3) the KGK principal will contribute her share at a later date; (4) the parties intend an equal (50/50) economic split once both have funded. Every bracketed item is highlighted in yellow. Open items requiring client input or attorney judgment:"));
children.push(
  bullet("Entity choice (Section 1.1): unincorporated joint venture vs. single-purpose LLC. Confirm state of organization and whether either Member is itself an LLC or corporation, and the full legal names as registered."),
  bullet("Economic structure pending KGK’s payment (Sections 3.3, 7.2(b), 8.3): the draft gives KGK a 0% interest plus a Contribution Right until it pays. The alternative is a 50/50 interest from day one with KGK’s unpaid share treated as a debt. Select one and conform Sections 7.2(b) and 8.3."),
  bullet("Deferred Contribution amount, deadline, whether interest accrues on Teper Holdings’ carry, and whether installments are permitted (Sections 3.2, Schedules B and C)."),
  bullet("Consequences of non-payment (Section 3.4): confirm the Members are comfortable with termination of KGK’s Contribution Right as a remedy, and whether a cure period should be added."),
  bullet("Custody and title (Section 4.2): confirm who physically holds the Cards and in whose name the Whatnot and PSA accounts are held. Consider a vault (e.g., PSA Vault) if the Cards are high-value."),
  bullet("Insurance (Section 4.4): confirm whether collectibles coverage exists or will be purchased, and whether the obligation should be firm or “commercially reasonable efforts.”"),
  bullet("Major Decision expense threshold (Section 5.3(g)) and whether the Managing Member should receive a management fee (Section 5.7)."),
  bullet("Target holding period and appraisal mechanism for the right of first offer (Sections 8.1, 8.2)."),
  bullet("Tax reporting: partnership return vs. IRC § 761(a) election out; collectibles gain rate (Section 9.3). Coordinate with the clients’ CPA."),
  bullet("Governing law, venue, and arbitration vs. litigation (Sections 13.1, 13.3)."),
  bullet("Representation disclosure (Section 13.10): confirm which party counsel represents and whether KGK will be advised to obtain separate counsel."),
  bullet("Schedule A: verify card descriptions, PSA certification numbers, serial numbering, purchase date, seller, and price against the PSA labels and the Whatnot order confirmation."),
);

// ---- build -------------------------------------------------------------------
const doc = new Document({
  creator: "Teper Holdings",
  title: "Joint Venture Operating Agreement — Teper Holdings / Group KGK Inc.",
  styles: {
    default: { document: { run: { font: FONT, size: BODY } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true }, paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
    ],
  },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "DRAFT — FOR ATTORNEY REVIEW — PRIVILEGED & CONFIDENTIAL", font: FONT, size: 16, color: "808080" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Teper–KGK Card Venture — Joint Venture Operating Agreement — Page ", font: FONT, size: 16, color: "808080" }),
                 new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "808080" }),
                 new TextRun({ text: " of ", font: FONT, size: 16, color: "808080" }),
                 new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: "808080" })] })] }) },
    children,
  }],
});

// docx-js emits <w:highlightCs/> after <w:highlight/>, which strict OOXML validators reject.
// Word ignores it; strip it so the file validates cleanly.
const JSZip = require("jszip");
Packer.toBuffer(doc)
  .then(buf => JSZip.loadAsync(buf))
  .then(async zip => {
    const xml = await zip.file("word/document.xml").async("string");
    zip.file("word/document.xml", xml.replace(/<w:highlightCs[^>]*\/>/g, ""));
    return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  })
  .then(buf => { fs.writeFileSync(OUT, buf); console.log("Wrote", OUT); });
