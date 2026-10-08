# Answer key: seeded issues in the sample plans

Scenario: Atlas Platform Migration. Four teams, four plans, deliberately different formats.

| File | Team | Date format | Dependencies expressed as |
|---|---|---|---|
| 1_core_team_plan.xlsx | Atlas Core | yyyy-mm-dd, Start/End | Task IDs (plus one vendor ID) |
| 2_vendor_team_plan.xlsx | CloudBridge (vendor) | dd-mmm-yyyy, Start/Finish | Mix of IDs and free text |
| 3_partner_a_payments_integration_plan.xlsx | PayNet Integration (Partner A) | mm/dd/yyyy + duration, **no end date** | IDs and free text |
| 4_partner_b_data_security_plan.xlsx | Data & Security (Partner B) | "Oct 19" text, **no year** | Free text only, **no IDs** |

Partner A end dates (working days, computed): PA-1 Oct 9, PA-2 Oct 23, PA-3 Nov 13, PA-4 Nov 27, PA-5 Dec 11, PA-6 Nov 20, PA-7 Dec 9, PA-8 Dec 18.

## Cross-team dependencies (the tool should surface these)

| # | Upstream | Downstream | Declared where |
|---|---|---|---|
| D1 | C-03 Provision cluster (ends Oct 30) | V-03 Install agents | Vendor, free text |
| D2 | C-06 Rate limiting design (ends Oct 23) | PA-2 Gateway API contract | Partner A, in notes only |
| D3 | V-05 Migration scripts (ends Nov 20) | C-07 Dry run | Core, by ID |
| D4 | C-08 / staging stability (ends Nov 20) | PB-4 Staging freeze (Nov 16) | Partner B, free text |
| D5 | V-03 / V-04 Monitoring agents + dashboards | PB-3 Ingestion pipeline | Partner B only. **Vendor plan has no "log forwarding" deliverable (gap)** |
| D6 | PA-4 Integration testing (ends Nov 27) | C-09 Load testing | Core, in notes only |
| D7 | Core refund endpoint timeout fix (no task in core plan) | PA-6 Refund validation | Partner A, free text. **No matching task to link to** |
| D8 | PB-7 Compliance sign-off, PA-5 Processor certification | C-11 Go/No-Go, C-12 Cutover | **Not declared in core plan (missing gates)** |

## Timeline conflicts

| # | Conflict | Size |
|---|---|---|
| T1 | Vendor assumes cluster ready Oct 23, Core delivers Oct 30. V-03 starts Oct 26 | 7 days |
| T2 | C-07 dry run starts Nov 16, V-05 scripts finish Nov 20. V-06 also starts before V-05 finishes | 4 days |
| T3 | **Cutover date mismatch**: Core and Partner B say Dec 7. Vendor and Partner A say Dec 14 | 7 days |
| T4 | PA-5 processor certification ends Dec 11, after Core cutover Dec 7 | 4 days |
| T5 | PB-6 remediation (Dec 7-11) and PB-7 compliance sign-off (Dec 11) fall after Go/No-Go (Dec 4) and cutover (Dec 7) | 4 days |
| T6 | C-09 assumes PayNet done Nov 20, PA-4 ends Nov 27 | 7 days |
| T7 | C-09 spans Nov 23-27 but includes US Thanksgiving (Nov 26-27), so only 3 working days | 2 days lost |
| T8 | PB-3 pipeline starts Nov 2, but V-03 agents are not installed until Nov 6 (dashboards Nov 20) | 4+ days |
| T9 | PB-4 staging freeze Nov 16, while C-07 and C-08 are active in staging Nov 16-20 | 4 days |
| T10 | PA-2 finishes Oct 23, the same day C-06 delivers the rate limits it needs | 0 float |
| T11 | PA-7 readiness review starts Dec 7 but depends on PA-5 (ends Dec 11) | 4 days |
| T12 | V-08 hypercare (Dec 14-23) overlaps vendor change freeze starting Dec 18 | 6 days |

## Resource conflict

- R1: Marcus Chen is requested by PA-4 (~50%), PA-6, and PB-5 (~50%) during Nov 16-27, while also supporting C-08 regression and owning C-05/C-12. Over-allocated.

## Risks and status mismatches

- S1: Vendor reports **Green** on V-03 even though its stated assumption (cluster by Oct 23) conflicts with Core's Oct 30.
- S2: V-05 is Amber and its slip directly threatens C-07, C-09, and the Dec 7 cutover.
- S3: PB-3 is At Risk, but the root cause (vendor log format) is not on any other team's plan.
- S4: Partner B reports every security task On Track despite T5 (sign-off after cutover).
- S5: The refund endpoint timeout appears in Core notes and the Partner A dependency but has no owned task or end date in any plan.
- S6: Core's critical path (C-03 to C-05 to C-08 to C-09 to C-10 to C-11 to C-12) has no float against T1, T2, T6, T7.

## Addendum: what plan_synthesizer.py finds

Found by the rules alone (no API call): D1, D2, D3, D5 (low confidence), D6 (low confidence), D7 (as an unresolved dependency), D8; T1-T8, T10-T12; R1; S1-S5.

Needs `--use-claude`: D4 / T9 (PB-4 "Atlas staging stable" has no keyword overlap with any Core task name, so the rules report it as unresolved; Claude can map it to C-08).

Two real conflicts I did not intentionally seed, which the tool also found inside the Core plan:
- T13: C-03 (cluster) starts Oct 12 but depends on C-02 (architecture sign-off), which ends Oct 16.
- T14: C-06 (rate limiting design) starts Oct 12 with the same C-02 dependency.
