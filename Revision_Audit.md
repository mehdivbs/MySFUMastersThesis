# Examiner-Feedback Revision Audit

**Audit date:** 2026-10-20
**Current manuscript:** `main.tex`, commit `1652864` (`Revise thesis based on reviewer feedback`)
**Pre-revision comparison:** parent commit `d31f3be`

## Scope and Evidence Available

The supplied workspace contains the thesis source, bibliography, SFU class file, figures, and presentation artifacts. Git history provides a pre-revision parent version. No separate examiner-report file, alternate thesis draft, change notes, or Google Drive folder is available in this workspace. The examiner concerns below are extracted from the review text supplied in the conversation; this audit cannot verify whether that text is the complete original report. The workspace has no ICADOS implementation repository or source code apart from screenshots embedded as figures, so implementation claims cannot be independently checked against executable software.

## File Inventory

| File or group | Role | Audit finding |
|---|---|---|
| `main.tex` | SFU thesis manuscript | Authoritative editable manuscript currently present; includes the prior reviewer-response edits. |
| `references.bib` | Bibliography | Bibliography used by `main.tex`; recent edits add PRISMA-ScR and Zaimoglu et al. references and repair a Derbel key. |
| `sfuthesis.cls` | SFU thesis format | Required document class; preserve unchanged unless compilation proves a necessary fix. |
| `Figures/` | Thesis figures and prototype screenshots | Relevant visual artifacts are present. Their underlying source files and system code are not present. |
| `Overleaf_Presentation/`, `generate_presentation.py`, `Create_Thesis_Presentation.ps1`, `Build_Thesis_Presentation.ps1` | Separate presentation work | Not part of the thesis evidence or an alternate thesis draft. |
| Git history | Version evidence | `1652864` is the current committed revision; `d31f3be` is the parent snapshot available for comparison. The working tree was clean at audit start. |
| Examiner report / Google Drive / revision notes | Requested source material | Not present or accessible in this workspace. The conversation contains reviewer comments and a later revision protocol, but these are not a separately verifiable report. |

## Thesis Argument Map

1. **Practical problem:** Data-driven tools can inform decisions, but the thesis argues their use should account for decision tasks, data, causal understanding, and decision-maker characteristics.
2. **Research gap as currently framed:** The manuscript argues that tools are fragmented and insufficiently aligned with human reasoning and decision stages. This gap is too broad until compared more systematically with prior work, including Zaimoglu et al. (2023) and Pratt et al. (2023).
3. **Questions:** RQ1 identifies technique classes in a selected corpus; RQ2 describes applications and limitations; RQ3 synthesizes human causal reasoning; RQ4 proposes an integrated process/framework.
4. **Evidence and methods:** RQ1 uses one query in Web of Science and PubMed, caps review at the top 100 results per database, and reports 72 selected papers. RQ2 is a selective narrative analysis without a documented example-selection rule or common quality appraisal. RQ3 is a narrative synthesis without recorded search details.
5. **Synthesis:** The manuscript groups techniques into six classes and maps illustrative applications to Simon's Intelligence, Design, and Choice stages. The evidence supports a descriptive, sample-bounded synthesis, not field-wide prevalence or comparative effectiveness.
6. **Framework:** Orchestration maps tool families to activities/stages; integration selects, customizes, and combines tools with task-, data-, and user-relevant considerations. This distinction is present but terminology and claims need a consistency pass.
7. **Prototype:** ICADOS is described through architecture, UI figures, and some source-code screenshots. The thesis itself says REST APIs were not implemented. No software repository is available here to verify implementation completeness; no full SCI workflow was applied to a real task.
8. **Conclusion boundary:** The framework and prototype remain conceptual/proof-of-concept artifacts. No user study or real-world evaluation establishes effectiveness, usability, confidence, or organizational impact.

## Reviewer-Concern Matrix

Status definitions: **1** fully addressed and integrated; **2** substantially addressed but needs refinement; **3** partially addressed; **4** mentioned but not substantively resolved; **5** not addressed; **6** cannot be verified from available evidence.

| # | Concern and location | Existing revision and adequacy | Remaining issue / next action | Status / priority |
|---|---|---|---|---|
| 1 | Vague “tons of these routine decisions” (Introduction, Background) | A clarifying paragraph was appended, but the original wording remains. | Rewrite the opening paragraph in place; remove colloquial wording and avoid claims about prevalence. | 3 / Medium |
| 2 | Meaning of the $250 million estimate (Introduction) | Added that it refers to wages and about 530,000 manager-days at a typical Fortune 500 company. | Integrate adjacent to the statistic and verify the exact source framing; distinguish survey estimate from measured savings. | 2 / Medium |
| 3 | Unsupported claim that human traits/reasoning were “generally neglected” (Introduction) | A paragraph retracts the claim and cites existing human-factors work. | The original claim remains in the preceding long paragraph, so the text contradicts itself. Rewrite the gap claim in place and align Chapter 2's “absence” language. | 3 / High |
| 4 | Aristotle-to-Kant discussion lacks relevance (Chapter 2, Decision Theory) | A transition says it is philosophical context, not empirical support. | The historical passage remains lengthy and weakly connected. Shorten it or explicitly connect only the relevant distinction to later decision theory. | 2 / Low |
| 5 | RQ1 needs an appropriate review structure | RQ1 was changed to ask which classes are represented in the sample; the search is called exploratory. | This accurately limits the claim but does not provide the stronger review structure the question merits. A new scoping review would require a documented protocol and fresh screening. | 4 / Critical |
| 6 | Database coverage and rationale (WoS/PubMed; ACM/IEEE/Scopus/AIS) | The omissions and broad database rationale are acknowledged. | No supplementary searches were performed; the corpus remains incomplete for HCI, computing, engineering, and information systems. Search and report a transparent supplementary review if feasible; otherwise retain as limitation. | 4 / Critical |
| 7 | PRISMA-ScR is reporting guidance, not the review method | The distinction is now stated and a reference added. | The thesis does not claim PRISMA-ScR compliance, appropriately; a future review should use a protocol and report verified flow counts. | 2 / High |
| 8 | Single query, top 100 per database, no search date | These constraints and missing search records are acknowledged in methods and limitations. | No historical date, retrieval counts, duplicate handling, or full screening flow can be reconstructed from current files. Do not invent these; disclose unavailable records. | 4 / Critical |
| 9 | Zaimoglu, Pratt, and Fisher (2023) and contribution relative to it | The article and Pratt et al. (2023) are cited in a new comparison paragraph. | Comparison is brief and placed at the start of the framework chapter, not developed alongside the literature review/design rationale. Verify the works and compare only supported dimensions. | 3 / High |
| 10 | RQ2 over-positive, quote-driven synthesis | The revision adds a selective-sample caveat and general class-level risks; some claims were narrowed. | Risks are partly appended instead of integrated into each tool-family argument. Many positive claims/quotes remain, and the RQ2 introduction still says “thoroughly analyzed” every coded study. | 3 / Critical |
| 11 | About 60 identified papers not discussed in depth | The text now says it uses selected examples and lacks a documented selection rule. | The identified corpus is not listed in a reproducible inventory; the exact papers and rationale cannot be recovered from available records. Preserve this as a limitation; do not invent post hoc criteria. | 4 / High |
| 12 | ML example may mainly be Miao's communication-engineering study | The revision now calls this one domain-specific example. | Verify the cited paper's actual methods/results and distinguish them from broader ML claims. Other unsupported generalizations remain. | 2 / High |
| 13 | Tool-class strengths and weaknesses promised for RQ2 | Several risks are named for ML, statistics, NLP, visualization, optimization, and simulation. | Coverage is uneven and mostly generic; link each limitation to evidence and decision conditions, distinguish potential from reported outcomes, and connect findings to design choices. | 3 / High |
| 14 | Only a subset of visualization/ML literature discussed | The absence of a documented sampling rule is acknowledged. | No verified selection process can be offered retrospectively. Clarify corpus coding versus illustrative examples and explicitly limit synthesis claims. | 4 / High |
| 15 | MOO stakeholder selection versus automatic choice | RQ2 summary now says Pareto alternatives are selected by stakeholders unless an explicit rule is supplied. | Framework mapping still says single or non-dominated solutions are “autonomously selected with minimum human intervention”; reconcile all Chapter 4/5 statements and distinguish prior, posterior, and interactive preferences. | 3 / Critical |
| 16 | RQ3 literature identification and repeated handbook citations | Chapter 3 now says a reproducible search was not documented and repeated handbook citations are not independent evidence. | The actual source-finding route cannot be verified; no additional empirical literature synthesis has been conducted. State this as unresolved and add verified sources only if further research is undertaken. | 4 / High |
| 17 | Expert causal belief versus empirical data | A new paragraph recommends checking assumptions, data quality, alternative explanations, and sensitivity. | It is unsourced and does not fully distinguish weak/biased data, contextual expert knowledge, and errors in both models. Develop a cautious synthesis with verified evidence or retain as a conceptual limitation. | 3 / High |
| 18 | RQ4 design rationale, coherence, and build-test | The abstract and conclusion now state that there was no empirical evaluation. | The framework's design choices are not consistently traced from evidence to principle to workflow to architecture; no build-test evaluation can be claimed. Improve traceability and mark untested workflows. | 3 / Critical |
| 19 | Abstract overstates practical effectiveness | The abstract now says the framework/prototype were not evaluated and makes no effectiveness claim. | Other Introduction and framework passages still use “effective,” “enhance,” and generalized benefit language. A thesis-wide claim audit is required. | 2 / High |
| 20 | SCI not performed; rules from limited sources; real-world impact unknown | The limitations now state that SCI was not completed for a real task and that impact is unknown. | Chapter 4 still describes systematic prescriptions and adaptive automation in stronger terms than the evidence supports. Align those descriptions and the system chapter. | 3 / Critical |
| 21 | Limitations omit review weaknesses and missing user study | RQ1/RQ3 search limitations and absence of evaluation were added to Chapter 6. | Add the inability to verify implementation from source code, uneven evidence quality, and context dependence; ensure no contradiction with the contribution and framework chapters. | 2 / High |
| 22 | Implementation status and unimplemented APIs | Chapter 5 explicitly says simulation REST APIs were not implemented. | Other passages describe the architecture as integrating tools and provide user workflows as if operational. No executable ICADOS repository is available to verify which features work. Distinguish conceptual, pictured, coded, and unimplemented features throughout. | 6 / Critical |

## Audit Conclusion

The current revision is a useful corrective pass, not a complete response to the examiner. It should be retained as the baseline, but several additions need integration and cross-chapter consistency work. The most urgent fixes are the contradictory Introduction wording, the mismatch between methods and RQ1 prevalence claims, the conflict between stakeholder choice and autonomous optimization, and the gap between conceptual/UI descriptions and unimplemented APIs. Historical search dates, screening counts, study-selection rules, source-code functionality, and empirical outcomes cannot be reconstructed from the available workspace and must not be fabricated.

A full revised manuscript can improve accuracy and integration, but it cannot convert the original searches into a systematic review or supply the missing user evaluation. Those remain explicit unresolved limitations unless new, transparently documented research is performed.
