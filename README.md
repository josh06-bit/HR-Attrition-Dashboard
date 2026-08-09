# HR-Attrition-Dashboard — Business Questions & Findings

An automated Excel dashboard built entirely in VBA over the IBM HR Employee Attrition dataset (1,470 employees × 35 fields). One macro — BuildHRDashboard — cleans the data, builds seven PivotTables, five charts, six KPI cards and two connected slicers from a raw CSV in about 30 seconds.

<img width="1682" height="690" alt="image" src="https://github.com/user-attachments/assets/e1fdbffd-b259-434f-9cf2-9ad385a09761" />


This document answers the ten questions an HR director would actually ask of it.

********

| Baseline KPIs |
| ------ | ----- |
| Metric | Value |
| Total employees | 1,470 |
| Employees exited | 237 |
| Attrition rate | 16.1% |
| Avg monthly income | $6,503 |
| Avg tenure | 7.0 years |
| Avg age | 36.9 years |

Every figure below is computed from the same source table and is reproducible by filtering the dashboard slicers.

 
## Q1. What is our overall attrition rate, and should we be worried?
 
**Answer.** 237 of 1,470 employees have left — a rate of **16.1%**.
 
**Insight.** On its own the headline number is unremarkable; it sits close to the cross-industry
average and would pass unnoticed in a board pack. The problem is that it is an *average of wildly
uneven groups*. Sub-populations inside this workforce range from 2.5% to 52.6%. Reporting 16.1%
alone conceals the fact that roughly one third of all exits come from a group representing barely
a tenth of the headcount.
 
**Recommendation.** Stop managing to the headline rate. Report attrition segmented by job level,
overtime status and tenure band as the standing monthly view, and set separate thresholds per
segment. A single company-wide target will always be met by the low-risk majority while the
high-risk pockets keep bleeding.
 
---
 
## Q2. Which department is losing the most people?
 
**Answer.** By *rate*, Sales is worst at **20.6%** (92 of 446), followed by Human Resources at
**19.0%** (12 of 63) and Research & Development at **13.8%** (133 of 961). By *volume*, the order
inverts: R&D accounts for **56%** of all exits, Sales 39%, HR 5%.
 
**Insight.** This is the classic rate-versus-volume trap, and it is why the dashboard charts
headcount and exits side by side. Sales has the sharper retention problem per employee, but R&D
loses the most actual people and therefore the most institutional knowledge. HR's 19.0% rests on
just 12 departures out of 63 staff — two or three more exits would swing that figure by five
points, so it should not drive decisions on its own.
 
**Recommendation.** Run two different interventions. Sales needs a structural fix aimed at the
role design (see Q4). R&D needs a volume-focused retention programme targeted at its two largest
job families. Treat the HR department's rate as directional only and review it quarterly rather
than monthly, given the small base.
 
---
 
## Q3. Which age group is most at risk?
 
**Answer.** Attrition falls steeply with age. The 18–25 band leaves at **35.8%** (44 of 123),
26–35 at **19.1%**, 36–45 at **9.2%**, 46–55 at **11.5%**, and 56+ at **17.0%**.
 
**Insight.** More than one in three employees under 26 leaves — over twice the company average.
The curve is U-shaped: risk collapses through the mid-career years and lifts again at 56+, though
that final uptick covers only 47 people and is likely retirement rather than regretted attrition.
The under-26 group is also 81% job level 1 with average tenure of 2.5 years, so age here is partly
a proxy for seniority and tenure rather than an independent driver.
 
**Recommendation.** Build a structured first-two-years programme for early-career hires —
assigned mentor, defined 24-month progression path, and a check-in at months 3, 6 and 12. Track
the under-26 cohort as its own retention metric. Separate voluntary from retirement exits in the
56+ band before treating it as a problem at all.
 
---
 
## Q4. Which job roles are the biggest flight risks?
 
**Answer.** Sales Representative leads at **39.8%** (33 of 83), then Laboratory Technician
**23.9%** (62 of 259), Human Resources **23.1%**, Sales Executive **17.5%** and Research Scientist
**16.1%**. At the safe end: Research Director **2.5%**, Manager **4.9%**, Healthcare
Representative and Manufacturing Director both **6.9%**.
 
**Insight.** The spread is 16-fold from safest to riskiest role — by far the widest of any
dimension on the dashboard, which is why Job Role is one of the two slicers. Sales Representatives
average $2,626 monthly income against a company average of $6,503, are 30.4 years old and hold 2.9
years of tenure: young, junior and the lowest-paid role in the business. By raw count, though,
Laboratory Technician (62 exits) and Sales Executive (57) each lose more people than Sales
Representative does.
 
**Recommendation.** Prioritise Laboratory Technician and Sales Executive for volume impact — those
two roles alone account for half of all exits. Treat Sales Representative as a role-design problem
rather than a people problem: benchmark its compensation externally and define a visible promotion
path into Sales Executive, since the current structure reads as a dead end.
 
---
 
## Q5. Does overtime actually drive attrition, or is that just folklore?
 
**Answer.** Employees working overtime leave at **30.5%** (127 of 416) versus **10.4%** (110 of
1,054) for those who do not — **2.9× higher**.
 
**Insight.** This is the single strongest binary split in the dataset. Overtime employees make up
28% of the workforce but contribute 54% of all departures. The relationship is correlational —
this dataset is a snapshot with no dates, so it cannot prove that overtime *causes* exits rather
than both being symptoms of understaffed teams — but the effect size is large enough that the
distinction matters less than the action does.
 
**Recommendation.** Audit overtime by team, not by individual, and treat any team with sustained
overtime as a headcount-planning signal rather than a performance one. Set a monthly overtime
ceiling with an escalation trigger. Given the size of this effect, overtime status is the most
useful single field for a predictive flight-risk flag.
 
---
 
## Q6. Is pay the reason people leave?
 
**Answer.** Partly. Leavers earn a median of **$3,202** per month against **$5,204** for those who
stay. Split into income quartiles, the bottom quartile leaves at **29.3%** and the top at
**10.3%**.
 
**Insight.** Pay clearly correlates with retention, but the effect concentrates almost entirely in
the lowest quartile — the gap between Q2 (14.2%), Q3 (10.6%) and Q4 (10.3%) is comparatively
narrow. That pattern says the issue is a floor problem, not a general competitiveness problem.
Above roughly the 25th percentile, more money buys very little additional retention. Income is
also heavily confounded with job level and role, so it should not be read as an independent lever.
 
**Recommendation.** Do not fund an across-the-board raise; the data does not support a return on
it. Commission an external benchmark for the bottom income quartile specifically, with Sales
Representative and Laboratory Technician first. Model the cost of lifting that floor against
replacement cost for the 108 exits that quartile produced.
 
---
 
## Q7. When do people leave — how long do they last?
 
**Answer.** Attrition falls sharply with tenure: **34.9%** in the first year (75 of 215), 18.4% at
2–3 years, 12.8% at 4–6, 12.4% at 7–10, and **8.1%** beyond 10 years. Median tenure at exit is
**3 years** versus 6 for those who stay.
 
**Insight.** More than one in three first-year employees leaves, and nearly a third of all exits
occur within the first three years. This is the strongest argument in the whole dataset that the
problem is front-loaded: onboarding, role clarity and early-career progression, not long-service
disengagement. Anyone who survives past year four is comparatively stable.
 
**Recommendation.** Concentrate retention spend in the first 24 months, where it will reach the
highest-risk population. Introduce a 90-day and 12-month structured review, and — because this
dataset has no hire or exit dates — begin capturing them so that cohort survival can actually be
tracked over time. That single schema change would unlock genuine trend analysis.
 
---
 
## Q8. Does seniority protect people?
 
**Answer.** Strongly. Job level 1 leaves at **26.3%** (143 of 543), level 2 at 9.7%, level 3 at
14.7%, level 4 at **4.7%** and level 5 at 7.2%.
 
**Insight.** Level 1 holds 37% of the workforce but produces **60% of all exits**. The decline is
not perfectly monotonic — level 3 sits above level 2, likely reflecting mid-career employees who
are marketable enough to move and not yet senior enough to be locked in — but the overall gradient
is unmistakable. Junior status is the common thread linking the age, tenure, income and role
findings above; they are largely four views of the same underlying population.
 
**Recommendation.** Make level 1 → level 2 progression an explicit, published, time-bound path,
since crossing that threshold is associated with a two-thirds drop in exit risk. Investigate the
level 3 anomaly separately with qualitative exit interview data — the quantitative data flags it
but cannot explain it.
 
---
 
## Q9. Does educational background predict who leaves?
 
**Answer.** Weakly. Human Resources graduates leave at **25.9%** (7 of 27), Technical Degree at
**24.2%** (32 of 132) and Marketing at **22.0%** (35 of 159), against Life Sciences 14.7%, Medical
13.6% and Other 13.4%.
 
**Insight.** This is the weakest signal on the dashboard and it is included deliberately, to show
where *not* to invest. The Human Resources field covers only 27 people, so its rate rests on seven
individuals and carries almost no statistical weight. The apparent pattern also largely dissolves
on inspection: Technical Degree and Marketing graduates cluster into the junior, lower-paid roles
already identified in Q4, which means education field is mostly restating role rather than adding
information.
 
**Recommendation.** Do not build hiring policy on this dimension — screening candidates by
educational background here would risk discriminatory practice while adding negligible predictive
value over job role. Use the chart for descriptive workforce composition only.
 
---
 
## Q10. If we could only act on one group, who should it be?
 
**Answer.** Employees at job level 1 who work overtime leave at **52.6%** (82 of 156). The same
job level *without* overtime leaves at 15.8%. A second overlapping profile — single employees
working overtime — leaves at **49.6%** (65 of 131).
 
**Insight.** This is the compound finding, and it is what the two connected slicers exist to
surface. That 156-person group is **10.6% of the workforce but produces 34.6% of all exits**. The
two risk factors do not merely add; the combined rate is roughly double what job level 1 alone
would predict. It is also the most actionable segment on the dashboard, because overtime is a
management-controllable variable in a way that age, tenure and marital status are not.
 
**Recommendation.** Make this the pilot cohort. Cap overtime for level 1 employees for two
quarters, pair it with the early-tenure check-in programme from Q7, and measure the cohort's exit
rate against a matched group. It is a small, well-defined population, the intervention is cheap,
and the effect size is large enough to be measurable within a year.
