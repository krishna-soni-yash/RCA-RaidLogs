# RAID log numbering

New items read the current `RaidLogID` values from SharePoint, filtered by
`SelectType`, and save the maximum numeric value plus one. Each of Risk,
Opportunity, Issue, Assumption, Dependency, and Constraints has an independent
sequence. Empty types start at 1. Browser filters and table pagination have no
effect on allocation. The query follows every page and aborts if any page fails.

`RaidLogID` must be a writable Number or Single line of text column. Existing
text prefixes and zero padding are preserved (for example, `RISK-0099` becomes
`RISK-0100`). Within a type, nonempty IDs must use a consistent prefix and end in
a nonnegative integer; invalid values stop creation. Empty values are ignored.
For a type with no existing IDs, text IDs also start at `1` without a prefix.

One new Risk gets one number shared by its Mitigation and Contingency records.
Adding an action to an existing Risk reuses that Risk's existing ID. Edits do not
rewrite IDs or allocate numbers. Existing records are not renumbered. Deleting
the highest ID allows its number to be reused under the maximum-existing-plus-one
rule.

## Deployment assumptions and concurrency

The allocating user needs read access to all items of the selected type; item-level
visibility restrictions would hide part of the sequence. For large lists, index
`SelectType` and validate the query against the site's list thresholds. Any
external automation that assigns `RaidLogID` must be reconciled before deploying
this writer.

The current maximum lookup and item creation are separate SharePoint operations.
The form prevents repeated Save clicks, but different users can still allocate
the same number concurrently. Full cross-user protection requires a shared
server-side coordination mechanism covering both lookup and creation; it is not
provided by this browser-only implementation. No supporting SharePoint list is
created automatically.

## Verification

Run `npm run test:raid-ids` for the isolated numbering and mocked service tests,
and `npm run build` for TypeScript, lint, and bundling. The automated tests do not
connect to a live SharePoint site.
