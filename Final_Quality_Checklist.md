# Final Quality-Control Checklist

**Manuscript under review:** `main_revised.tex`
**Status key:** Addressed = revised text provides a clear and integrated response; Partial = the text acknowledges or narrows the issue but evidence/research remains insufficient; Unresolved = cannot be resolved from current materials; Unverifiable = required source material is unavailable.

| # | Examiner concern | Status after this revision | Verification / remaining limit |
|---|---|---|---|
| 1 | Clarify “tons of these routine decisions” | Addressed | Opening paragraph rewritten in academic language; removed colloquial phrase. |
| 2 | Explain the $250 million estimate | Addressed | Contextualized as a typical Fortune 500 wage estimate and not a projected DSS saving. |
| 3 | Correct broad claim that human reasoning was neglected | Addressed | Original phrasing removed; gap narrowed to organization around stages/tasks and design factors. |
| 4 | Connect or remove Aristotle-to-Kant detour | Addressed | Replaced with a short transition to causal understanding and action choice. |
| 5 | Use an appropriate RQ1 review structure | Partial | RQ1 and methods accurately describe an exploratory sample; no scoping review was retrospectively claimed. A proper scoping review remains future work. |
| 6 | Justify and broaden database coverage | Partial | WoS/PubMed rationale and omissions are explicit; ACM, IEEE, Scopus, and AIS were not searched. |
| 7 | Characterize PRISMA-ScR accurately | Addressed | Identified as reporting guidance; no compliance claim. |
| 8 | Report search date, query, cap, flow, and screening | Partial | Query, database, top-100 cap, and informal screening are stated; absent historical date, counts, and duplicate records remain unavailable. |
| 9 | Discuss Zaimoglu, Pratt, and Fisher (2023) | Partial | Added and positioned relative to the thesis; the comparison is conceptual and not a full dimension-by-dimension review. |
| 10 | Make RQ2 critically analytical | Partial | Added/rewrote class-level limitations and bounded claims; selected study evidence has not been appraised systematically. |
| 11 | Explain why only some papers receive detailed discussion | Partial | Explicitly states examples were selected without a documented rule; cannot reconstruct an ex post rationale. |
| 12 | Attribute and bound the Miao ML example | Addressed | Framed as a single communication-engineering example, not broad evidence for all ML applications. |
| 13 | Cover strengths and weaknesses by tool class | Partial | Summary now names risks and conditions across classes; detailed source-supported critique remains uneven. |
| 14 | Explain visualization/ML example selection and uncited papers | Partial | Acknowledges selective treatment and lack of selection rule; a corpus-level study inventory and systematic re-analysis remain needed. |
| 15 | Reconcile MOO trade-off sets and automatic selection | Addressed | Pareto sets are distinguished from single solutions requiring explicit preferences/objectives; ICADOS optimization is marked as proposed. |
| 16 | Explain RQ3 source-finding process and handbook citations | Partial | Non-reproducible narrative synthesis and repeated handbook reliance are disclosed; original search history is unavailable. |
| 17 | Address conflicts between expert beliefs and data | Partial | Adds a balanced, cited conceptual procedure; it has not been operationalized or evaluated. |
| 18 | Position the framework's originality and design rationale | Partial | Contribution is narrowed relative to prior work; no empirical advantage or novelty claim is made. Traceability of every design choice remains incomplete. |
| 19 | Align the abstract and claims with no evaluation | Partial | Abstract, Introduction, Chapter 5 framing, and Conclusion now disclaim demonstrated effectiveness; some broad benefit language remains elsewhere and needs a final whole-document edit. |
| 20 | State whether SCI and the system were implemented | Partial | Chapter 5 labels use cases and UI as proposed, identifies unimplemented REST APIs, and states no evaluation; no implementation code was available to verify other components. |
| 21 | Put search weaknesses and missing user study in Limitations | Addressed | Search limits, selective synthesis, unimplemented task-level SCI, and absent evaluation are explicitly listed. |
| 22 | Verify code, APIs, and other implementation claims | Unverifiable | No ICADOS source repository is available; only descriptions and figure-level examples can be inspected. |

## Consistency and Technical Checks

- [x] Original `main.tex` preserved; revisions are in a separate `main_revised.tex`.
- [x] RQ1 no longer claims field-wide “most common” techniques.
- [x] RQ1, RQ2, and RQ3 evidence approaches are distinguished.
- [x] Orchestration and integration are separately defined; task/data/user properties are treated as design factors.
- [x] MOO Pareto generation and preference-based final selection are distinguished.
- [x] Abstract and Chapter 5 describe a proposal/proof-of-concept, not validated effectiveness.
- [x] New citation keys resolve to bibliography entries.
- [x] Editor source diagnostics report no errors in `main_revised.tex`.
- [ ] LaTeX compilation and PDF visual inspection completed. Not run: `pdflatex` is unavailable in the current terminal.
- [ ] Every citation and quotation verified against full primary sources. Not completed; see `Unresolved_Issues.md`.
- [ ] ICADOS implementation claims checked against source code. Not possible: source repository unavailable.
- [ ] Examiner report and Google Drive folder independently inventoried. Not possible: neither was present/accessibly linked in the workspace.

This checklist confirms textual revisions only. It does not certify a reproducible scoping review, complete software implementation, or empirical validation.
