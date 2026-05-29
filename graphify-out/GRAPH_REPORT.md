# Graph Report - .  (2026-05-29)

## Corpus Check
- 24 files · ~42,662 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 126 nodes · 197 edges · 18 communities (13 shown, 5 thin omitted)
- Extraction: 95% EXTRACTED · 5% INFERRED · 0% AMBIGUOUS · INFERRED: 9 edges (avg confidence: 0.87)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Import Logic & User Mapping|Import Logic & User Mapping]]
- [[_COMMUNITY_Export Functions & Rationale|Export Functions & Rationale]]
- [[_COMMUNITY_Corpus Metadata (Graphify)|Corpus Metadata (Graphify)]]
- [[_COMMUNITY_No-Comments Export Variant|No-Comments Export Variant]]
- [[_COMMUNITY_Auth Scopes & Admin Check|Auth Scopes & Admin Check]]
- [[_COMMUNITY_Security & Error Handling|Security & Error Handling]]
- [[_COMMUNITY_Graph API Protocol|Graph API Protocol]]
- [[_COMMUNITY_Use Case & Known Limits|Use Case & Known Limits]]
- [[_COMMUNITY_Claude Settings Permissions|Claude Settings Permissions]]
- [[_COMMUNITY_Integration Tests|Integration Tests]]
- [[_COMMUNITY_Export Unit Tests|Export Unit Tests]]
- [[_COMMUNITY_Claude Settings File|Claude Settings File]]
- [[_COMMUNITY_Community 16|Community 16]]
- [[_COMMUNITY_Community 17|Community 17]]

## God Nodes (most connected - your core abstractions)
1. `Import-PlanFromJson` - 10 edges
2. `Write-PlannerLog()` - 8 edges
3. `Write-PlannerLog()` - 8 edges
4. `Write-PlannerLog()` - 7 edges
5. `Write-PlannerLog (Export)` - 7 edges
6. `Resolve-UserId` - 7 edges
7. `Write-PlannerLog()` - 7 edges
8. `files` - 6 edges
9. `Export-PlanDetails` - 6 edges
10. `CLAUDE.md - Project Instructions` - 6 edges

## Surprising Connections (you probably didn't know these)
- `Write-PlannerLog (Export)` --semantically_similar_to--> `Write-PlannerLog (Import)`  [INFERRED] [semantically similar]
  Export-PlannerData.ps1 → Import-PlannerData.ps1
- `Invoke-GraphWithRetry` --conceptually_related_to--> `Rate Limiting and Retry Logic`  [INFERRED]
  Import-PlannerData.ps1 → CLAUDE.md
- `Test-SafePath (Export)` --semantically_similar_to--> `Test-SafePath (Import)`  [INFERRED] [semantically similar]
  Export-PlannerData.ps1 → Import-PlannerData.ps1
- `Connect-ToGraph (Export)` --semantically_similar_to--> `Connect-ToGraph (Import)`  [INFERRED] [semantically similar]
  Export-PlannerData.ps1 → Import-PlannerData.ps1
- `Export-PlanDetails` --implements--> `OData Paging with nextLink`  [EXTRACTED]
  Export-PlannerData.ps1 → CLAUDE.md

## Hyperedges (group relationships)
- **Export Pipeline: Group Selection â†’ Plan Loading â†’ Data Export â†’ Summary** — export_fn_get_allm365groups, export_fn_get_groupsbynames, export_fn_show_groupselectionmenu, export_fn_get_plansbygroupids, export_fn_get_alluserplans, export_fn_export_plandetails, export_fn_export_readablesummary [INFERRED 0.92]
- **Import Pipeline: JSON Parse â†’ Plan Create â†’ Bucket Create â†’ Task Create â†’ Detail Set** — import_fn_import_planfromjson, import_fn_invoke_graphwithretry, import_fn_resolve_userid, import_fn_add_errortotracker, import_var_errortracker [INFERRED 0.90]
- **Security Validation Pattern: Path Check + UNC Block + GUID Validation** — export_fn_test_safepath, import_fn_test_safepath, concept_unc_path_blocking, concept_path_validation [EXTRACTED 0.95]

## Communities (18 total, 5 thin omitted)

### Community 0 - "Import Logic & User Mapping"
Cohesion: 0.14
Nodes (18): Rationale: Client-side sorting instead of $orderby (Graph HTTP 400 bug), Rationale: DateTimeOffset.Parse instead of DateTime.Parse (locale-independent), Rationale: PSCustomObject for exportSummary instead of Hashtable (Measure-Object compatibility), Export-ReadableSummary, Get-AllM365Groups, Get-AllUserPlans, Get-GroupsByNames, Get-PlansByGroupIds (+10 more)

### Community 1 - "Export Functions & Rationale"
Cohesion: 0.14
Nodes (16): DryRun Mode, Exit Codes for Automation (0/1/2), O(nÂ²) Task Detail Lookup Anti-pattern (known future improvement), Rationale: User resolution caching to reduce API calls (90+ seconds saved per plan), Cross-Tenant User Mapping, User Resolution Strategy, TODO.md - Implementation Status, Add-ErrorToTracker (+8 more)

### Community 2 - "Corpus Metadata (Graphify)"
Cohesion: 0.22
Nodes (16): Add-ErrorToTracker(), Connect-ToGraph(), Import-PlanFromJson(), Invoke-GraphWithRetry(), Resolve-UserId(), Write-CacheStatistics(), Write-ErrorSummary(), Write-PlannerLog() (+8 more)

### Community 3 - "No-Comments Export Variant"
Cohesion: 0.15
Nodes (12): files, code, document, image, paper, video, graphifyignore_patterns, needs_graph (+4 more)

### Community 4 - "Auth Scopes & Admin Check"
Cohesion: 0.29
Nodes (9): Rationale: NOCOMMENTS variant kept in sync with main script, Connect-ToGraph(), Export-PlanDetails(), Export-ReadableSummary(), Get-AllM365Groups(), Get-AllUserPlans(), Get-GroupsByNames(), Get-PlansByGroupIds() (+1 more)

### Community 5 - "Security & Error Handling"
Cohesion: 0.22
Nodes (7): Graph Scopes for Check-Admin-Rechte.ps1, Graph Scopes for Export (Group.Read.All, Tasks.Read, Tasks.ReadWrite, User.Read, User.ReadBasic.All), Graph Scopes for Import (Group.ReadWrite.All, Tasks.ReadWrite), PowerShell ISE Compatibility (WAM disable), Rationale: throw instead of exit 1 for ISE compatibility, Connect-ToGraph (Export), Connect-ToGraph (Import)

### Community 6 - "Graph API Protocol"
Cohesion: 0.47
Nodes (6): ETag Handling for PATCH Operations, OData Paging with nextLink, Planner Data Model (Plan/Bucket/Task/Details), Rationale: ETag required for PATCH - must GET first, CLAUDE.md - Project Instructions, Export-PlanDetails

### Community 7 - "Use Case & Known Limits"
Cohesion: 0.40
Nodes (4): Known API Limitations (Comments, Attachments, History), License Migration Use Case, Rate Limiting and Retry Logic, WARNING.md - Disclaimer and Limitations

### Community 8 - "Claude Settings Permissions"
Cohesion: 0.83
Nodes (4): Path Validation Security (Test-SafePath), UNC Path Blocking Rationale, Test-SafePath (Export), Test-SafePath (Import)

## Knowledge Gaps
- **25 isolated node(s):** `code`, `document`, `paper`, `image`, `video` (+20 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **5 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `PowerShell ISE Compatibility (WAM disable)` connect `Security & Error Handling` to `Import Logic & User Mapping`?**
  _High betweenness centrality (0.041) - this node is a cross-community bridge._
- **Why does `Import-PlanFromJson` connect `Export Functions & Rationale` to `Corpus Metadata (Graphify)`, `Graph API Protocol`?**
  _High betweenness centrality (0.040) - this node is a cross-community bridge._
- **Why does `CLAUDE.md - Project Instructions` connect `Graph API Protocol` to `Import Logic & User Mapping`, `Export Functions & Rationale`, `Corpus Metadata (Graphify)`?**
  _High betweenness centrality (0.034) - this node is a cross-community bridge._
- **What connects `code`, `document`, `paper` to the rest of the system?**
  _25 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Import Logic & User Mapping` be split into smaller, more focused modules?**
  _Cohesion score 0.13768115942028986 - nodes in this community are weakly interconnected._
- **Should `Export Functions & Rationale` be split into smaller, more focused modules?**
  _Cohesion score 0.14035087719298245 - nodes in this community are weakly interconnected._