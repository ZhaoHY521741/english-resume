# Content rules for MNC applications

The layout is reusable; these rules are what make the resume survive a screener at a
foreign-owned company. They are separated from `layout-spec.md` because they are about
judgement, not millimetres.

## 1. What a foreign-owned employer expects to see

| Section | Included? | Why |
|---|---|---|
| Name | Yes | Centred, large. Write it family-name-first (`ZHAO Haiyang`) if your name is Chinese — it removes the "which is the surname" guess. |
| Contact line | Yes | City, phone with country code, email. |
| **Availability** | **Yes** | `Available: <start> – <end> \| <days>/week \| <cities>`. Screeners filter on this before reading anything else. Omitting it loses interviews. |
| PROFILE | Yes, first | 3–4 lines. The only paragraph most screeners read in full. |
| EDUCATION | Yes | Degree, institution, dates, relevant coursework, and the postgraduate-recommendation fact if it is a plus. |
| EXPERIENCE / INTERNSHIP | Yes | Real employers, newest first. |
| PROJECTS | Yes | Where a student shows ability. Lead with the project that matches the target role. |
| AWARDS | Yes, own section | Competitions and scholarships. Burying them inside a project bullet hides them. |
| SKILLS | Yes | Grouped by category; no self-rated percentages. |

## 2. What to delete — and why

| Delete | Reason |
|---|---|
| Photo | Standard practice for US/UK employers is no photo. Including one reads as "doesn't know the convention" and can trigger anti-discrimination policy. |
| Date of birth / age | Same. Age is a protected characteristic; supplying it invites a compliant employer to discard the file. |
| **Political affiliation** (e.g. "Member of CPC") | In a foreign-owned company this is at best noise and at worst a signal a screener does not want to evaluate. It is a *positive* for SOEs and a *liability* for MNCs. |
| Marital status, height, weight, ethnicity | Not requested, not used, can only hurt. |
| ID number | Never. |
| Self-rated skills ("Python: 8/10") | Unverifiable and meaningless across candidates. Show the artefact instead. |
| Hobbies with no signal ("reading, travel") | Either it demonstrates something relevant or it takes a line you could use. |
| "Objective: seeking a challenging position…" | Generic. Replace with one concrete sentence about what you build. |

## 3. Section order

`PROFILE → EDUCATION → EXPERIENCE → PROJECTS → AWARDS → SKILLS`

Two deviations from a common Chinese template are deliberate:

- **PROFILE belongs at the top, not the bottom.** A screener reads top-down and decides in
  seconds. A profile in the last section is never read.
- **AWARDS belongs in its own section.** A line like
  `Awards: National Third Prize in GMC (2026) and MathorCup (2025); Provincial Second Prize…`
  buried as a bullet inside a *different* project is the single most common waste of a
  strong credential — the reader can no longer tell which project it belongs to.

## 4. Writing rules

1. **Start every bullet with a past-tense verb**: Built, Integrated, Analyzed, Delivered,
   Reduced, Automated, Negotiated. Never start with "Responsible for" or a noun phrase.
2. **One claim per bullet, with a number.** If a bullet has no number, ask what changed and
   by how much. See `phrasing.md` for patterns.
3. **Say what you made, not what you used.** "Used Dify" is a tool name; "delivered a Q&A
   service that cut lookups from 20 min to 2 min" is evidence.
4. **No unverifiable adjectives.** Delete "显著", "greatly", "significantly" unless a number
   follows.
5. **Never invent numbers.** If the original only says "reduced errors", write
   `reduced factual errors on long-document Q&A (measured on a self-built 50-case eval set)`
   and leave the percentage out until you can measure it.

## 5. The one-page rule

A student resume is one page. When it overflows, cut in this order:

1. A bullet that repeats another bullet's claim.
2. The least relevant project (for an AI role, the market-research project goes before the
   Power BI one only if the role is analytical).
3. Coursework lists — keep the 5 courses closest to the target role.
4. Bullets that describe process ("participated in meetings") rather than output.

Shrinking type below 9 pt is the *last* resort; the builder warns when it happens.

## 6. ATS notes

Most MNC career portals parse the PDF before a human sees it.

- Keep the PDF text-based (this builder draws real text, not an image).
- Avoid text boxes, columns built from tables, and headers/footers — this template uses
  tab stops, not tables, so parsing stays clean.
- Spell out both the acronym and the expansion once: `Global Management Challenge (GMC)`.
- Use the job posting's own nouns for skills (e.g. "LLM", "RAG", "Power BI") — the parser
  matches literal strings.

## 7. Checking that the company is genuinely foreign-owned

"Boss is a foreigner" is the wrong test — heads of Microsoft China, SAP China and Amazon
China are all Chinese nationals. Use these instead:

| Test | Genuine MNC | In-name-only / agency |
|---|---|---|
| Ownership | Company registry shows a foreign legal-person sole proprietorship | Chinese shareholders behind a foreign-sounding name |
| Where the posting lives | On the parent company's own domain (`careers.microsoft.com`, `jobs.sap.com`) | On a job board as "a well-known foreign enterprise" |
| Language | English JD, English interview round, reporting line abroad | Chinese-only JD |
| Recruiter | Direct | "XX Human Resources Services Co., Ltd." reposting |

Apply only through: the company's own careers page, LinkedIn postings from the company
itself, and the university careers office. Skip any JD that will not name the employer.
