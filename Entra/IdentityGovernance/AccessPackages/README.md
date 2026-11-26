# Access package creation

[Create-AccessPackage.ps1](./Create-AccessPackage.ps1) creates one group-based
access package and request policy per CSV or Excel row. It reuses the mandatory
`-CatalogName` or creates a published catalog if no exact match exists.
Ambiguous catalog names stop the run; unpublished catalogs must be published
first. Existing packages with the same name in that catalog are skipped,
never overwritten. Duplicate names within an input file are processed once.

## Prerequisites and use

Connect to Microsoft Graph before running the script. Requires the
`Microsoft.Graph.Authentication` and `Microsoft.Graph.Groups` modules, appropriate Entra ID Governance
licensing, and an account permitted to create catalogs, packages and policies.
Adding a group to a catalog also requires ownership of that group or a role
that permits modifying it.

Policies exclude guests, avoiding the guest-governance `NoLinkedSubscription`
requirement encountered with guest-inclusive policies. Member-user licensing
requirements still apply. The script does not change billing configuration.

For delegated access:

```powershell
Connect-MgGraph -Scopes 'EntitlementManagement.ReadWrite.All','Group.Read.All','User.ReadBasic.All'

.\Create-AccessPackage.ps1 -CsvPath '.\access-packages.csv' `
    -CatalogName 'Business applications' `
    -FallbackApproverIds '55555555-5555-5555-5555-555555555555' `
    -DefaultAppOwnerIds '66666666-6666-6666-6666-666666666666' `
    -WhatIf
```

Replace the example IDs with real user/group object IDs. Remove `-WhatIf`
after reviewing the preview. It performs reads and validates approvers and
groups, but makes no Graph writes, even when the catalog does not yet exist.
Application authentication additionally needs `Group.ReadWrite.All` to add
groups to catalogs and `User.Read.All` to validate approvers. The script checks
delegated scopes; application permissions and resource-specific roles are
enforced by Graph. Hidden-membership groups may require `Member.Read.Hidden`.

If delegated permissions are missing from `Get-MgContext`, the script stops
before Graph requests and lists every missing permission category, accepted
alternatives, current scopes and a reconnect command. Administrator consent
may be required. Scopes do not grant Entra roles: activate any required PIM
role before reconnecting. A Graph 403 can still indicate missing role permissions
even when the scope check passes.

The script only manages cloud-assigned security groups and Microsoft 365
groups. Synchronized, dynamic, distribution, mail-enabled security and
role-assignable groups fail explicitly rather than creating unusable packages.
It adds the group's **Member** role, never its Owner role.
Roles are read from the catalog's `resourceRoles` endpoint, filtered by origin
system and resource ID; expanded resource role lists can be empty in Graph.

## Backfill existing group members

Assignment backfill is enabled by default (`$CreateAssignments = $true` near
the top of the script). Set that variable to `$false` in the script to disable
backfill. There is no `-CreateAssignments` argument.

Rerun the same input and catalog with `-WhatIf` to preview backfill without
writes. Review the proposed users and get approval before removing `-WhatIf`.
Limit the input to the packages you intend to backfill; missing packages are
still created in a live run.

A matching package is not recreated. The backfill verifies that its actual
resources match the input and that it has exactly one matching `<name> - Requests`
policy with a 365-day lifecycle. It submits one administrative request per
eligible user, skipping users with an existing non-expired assignment or a
pending add request. `Submitted` is not confirmation of delivery.

Only direct directory member users are eligible. For packages with multiple
groups, a user must already be a direct member of every group, since an
assignment grants all the package's roles. The backfill does not change group
membership directly, update an existing policy, or remove assignments.
The Graph connection must be able to read each member's `userType` (use
`User.Read.All` if basic reads do not return it). Run only one backfill at a
time and verify delivery in Entra before retrying uncertain requests.

## Excel input

Install `ImportExcel` once in the PowerShell environment you use to run the
script. Excel itself is not required.

### Draft packages from application discovery

The discovery workflow keeps access-review and package creation separate.
Run these commands from the `IdentityGovernance` directory after connecting:

```powershell
Connect-MgGraph -Scopes 'AccessReview.ReadWrite.All','Application.Read.All','Group.Read.All','User.Read.All','EntitlementManagement.ReadWrite.All','RoleManagement.Read.Directory','Directory.Read.All'

# One application:
.\AccessReviews\Create-AccessReview.ps1 -ServicePrincipalId '<enterprise application object ID>' `
    -ReviewerType ResourceOwner -WhatIf -DiscoveryOutputPath '.\discovery.json'

# Or multiple applications in a CSV with a ServicePrincipalId column:
.\AccessReviews\Create-AccessReview.ps1 -CsvPath '.\applications.csv' `
    -ReviewerType ResourceOwner -WhatIf -DiscoveryOutputPath '.\discovery.json'

.\AccessPackages\Export-DiscoveryAccessPackages.ps1 -DiscoveryPath '.\discovery.json' `
    -WorkbookPath '.\packages.xlsx'

# Alternatively, use one package per discovered group without application-assignment reads:
.\AccessPackages\Export-DiscoveryAccessPackages.ps1 -DiscoveryPath '.\discovery.json' `
    -WorkbookPath '.\packages-per-group.xlsx' -OnePackagePerGroup

# Edit the workbook, then preview and review every result before removing WhatIf.
.\AccessPackages\Create-AccessPackage.ps1 -InputPath '.\packages.xlsx' `
    -CatalogName 'Business applications' -FallbackApproverIds '<fallback user object ID>' -WhatIf
```

The first command produces the normal access-review creation preview and a
structured discovery file; it does not create access reviews. The CSV form
includes all service principals from its `ServicePrincipalId` rows in one
combined workbook; each package row retains its originating application ID.
Both files and the workbook require new filenames. Store them in a private
directory: they contain application and group information. Discovery records
direct-user counts but not user membership lists.

By default, each directly assigned root proposes one package with only the root
group as a Member resource. The package adds assigned users to the root only;
it does not add them to nested groups. Nested groups in the root's membership
hierarchy, and their existing members' transitive membership in the root, are
explained in the description. Shared groups are described under each root's
separate package, not merged into one package. The description lists direct
Entra enterprise-application assignments by group. Assignments on nested groups
apply to those groups' own members, not to users merely assigned to the root
through this package. An application assignment does not guarantee effective
access to the application.
Reading group application assignments requires delegated or application
`Directory.Read.All` (or `Directory.ReadWrite.All` or
`AppRoleAssignment.ReadWrite.All`), admin consent, and any required delegated
directory role.

With `-OnePackagePerGroup`, each unique discovered root or nested group has its
own package containing only that group, including groups with no direct users.
A group shared by roots or applications produces one package; the first
discovered application's ID is retained as its discovery source. Its suggested
package name uses that application's name, but nested-group discovery does not
prove that group is directly assigned to the application or that its members
have effective application access. The description identifies the group but
does not list its applications. This mode does not request group application
assignments and does not require `Directory.Read.All` for that lookup. Other
discovery, package-creation and backfill permissions still apply. Before
creation, the source application's discovery and group name are rechecked.

The `Topology` sheet shows each discovered application-assigned root and its
nested descendants, including candidates that are not ready. A group can appear
under multiple roots; this sheet is not a list of all applications the groups
affect. Check `Manual review` for failed or uncertain checks, correct suggested
names that violate naming rules, and complete approvers and business details.
The exporter suggests a `Role` when a root or nested descendant group name
contains a standalone or camel-case `Admin`, `Administrator`, or `Superuser`
term (including plurals and separated `Super User`). This name-based heuristic
does not verify privileges or detect every privileged group. Review and edit the
suggestion: the `Role` value is included in the published description. Unknown
values are left blank. Blocked rows are not imported; resolve them and
regenerate the workbook. Retain the discovery ID columns and `Discovery metadata`
sheet. Before creation, the script rechecks the generated description; changed
group names or, in the default mode, application assignments require a
regenerated workbook.

The export and creation recheck enumerate packages across catalogs and compare
their complete resource roles and scopes. A confirmed active exact match is
skipped; shared groups and partial overlaps with other packages do not alone
block new version 3 or single-group version 4 workbooks. A matching group set with extra roles, uncertain
activity, incomplete reads, or forbidden Graph reads still require manual
review. Policies without enough information to determine activity are not
assumed inactive. Older version 1 and 2 workbooks retain their original
shared-group safeguards and descriptions; regenerate them to use these modes.
To verify tenant-wide visibility, this workflow currently requires a delegated
connection with a **directly assigned active directory-wide** Identity Governance
Administrator or Global Administrator role, or an app-only connection whose
Microsoft Graph application permissions include `EntitlementManagement.Read.All`
or `EntitlementManagement.ReadWrite.All`. Application permissions require admin
consent. After granting a permission, reconnect so the current connection
includes it. Catalog-scoped roles alone cannot establish tenant-wide visibility.
The delegated role check also requires `RoleManagement.Read.Directory` (or a
supported higher permission) and a role allowed to read role assignments. Hidden
membership also needs `Member.Read.Hidden`. If visibility or any required Graph
read cannot be verified, candidates stay in `Manual review`. App-only package
creation still requires the separate write, group, user, directory and
access-review permissions needed by the creation and discovery steps.

On each package run the script refreshes discovery once, compares each row with
the current root and descendants, and rechecks package coverage before creating
that row. It does not backfill a skipped existing package in this workflow.
Default backfill still applies to newly created packages, and only direct members
of **every** included group qualify. The checks cannot prevent simultaneous
creation by another administrator; avoid concurrent runs. Large tenants may
take significant time because coverage is refreshed for each candidate.

```powershell
Install-Module ImportExcel -Scope CurrentUser
.\Create-AccessPackage.ps1 -InputPath '.\access-packages.xlsx' `
    -CatalogName 'Business applications' `
    -FallbackApproverIds '55555555-5555-5555-5555-555555555555' `
    -WhatIf
```

The script reads the first worksheet with headers in row 1:

| Column | Meaning |
| --- | --- |
| `Access Package Name` | The package name; the same naming rules as CSV apply. |
| `GroupId` | One or more Entra group object IDs separated by commas, semicolons, or line breaks. If the column exists, it must be populated for each package and `Access Group` is ignored. |
| `Access Group` | Used only by older workbooks without `GroupId`: one exact, unique group display name or one or more group object IDs. |
| `Description`, `Role`, `Brand`, `Region`, `Potential Teams` | Description text and optional details. |
| `Approver 1 List` | Required first approval stage. |
| `Approver 2`, `Approver 3` | Optional additional stages, in order. Stage 3 requires stage 2. Blank or `NA` omits a later stage. |

Each stage supports `<Employee Manager>`, `<Line Manager>` or `<Manager>`;
`<Resource Owner>` or `<Group Owner>`; and user object IDs, UPNs, or `mail`
addresses. Values can be combined with semicolons, commas or line breaks.
Tokens are case-insensitive, and angle brackets are optional for the named
rules. The manager may be selected in only one stage, with the supplied fallback
users. Each stage requires a decision from only one of its listed approvers.

Resource-owner selection reads the **current human owners of every listed
group** and captures them as named approvers in the policy. Ownership changes
after creation do not update this list. Each group must have at least one human
owner and all owner lookups must succeed; otherwise the row fails before writes.
Self-review is not an approval-stage option: Entra does not allow requesters to
approve their own requests. Access reviews are separate and remain disabled.

The script uses `Get-MgGroup` to fetch and validate each group by ID. The
preview shows the Entra display names and IDs. All listed groups are added as
Member roles to the same package. For multiple groups, the reported `UserCount` is the
largest individual transitive user count, not the sum.
Stage selection does not depend on group size. An address must match exactly one Entra user across UPN
and `mail`. For ambiguous addresses, use the user object ID. Display names
and secondary mail aliases aren't supported.

If the workbook cannot be opened, confirm it opens in desktop Excel and save
a local copy as **Excel Workbook (`.xlsx`)** if your organization's policy
allows it. Password-encrypted or rights-managed copies may not be readable
by the importer. A reader error stops the run before worksheet validation.

## CSV and naming

Start with [access-packages.example.csv](./access-packages.example.csv).
The comma-delimited CSV requires `GroupId` and `AppOwnerIds`, plus either `Name`
or both `Application` and `Role`. `Description`, `Application`, `Role`, `Region`
and `Comments` are optional columns:

| Column | Meaning |
| --- | --- |
| `Name` | Business-friendly application/service name and the access provided. Ignored when the row has both `Application` and `Role`. |
| `Description` | Explain what the requester can do with this access. |
| `Application` | Optional. With `Role`, the package name becomes `Application - Role`. |
| `Role` | Optional. With `Application`, the package name becomes `Application - Role`. |
| `Region` | Optional. Added to the description as `Region: ...`. |
| `Comments` | Optional. Added to the description as `Comments: ...`. |
| `GroupId` | One Entra group object ID, not its name. |
| `AppOwnerIds` | Explicit approver **user** object IDs separated by semicolons. Blank uses `-DefaultAppOwnerIds`. |

When a row has both `Application` and `Role`, that is the end-user package name
(any `Name` is ignored) and the naming rules below apply to it. The description
is `Description` followed by one line per remaining detail (`Region`, `Comments`,
and `Application`/`Role` if only one was supplied). Blank values are omitted.

Use the application/service name requesters recognize, describe the access
instead of the Entra group, explain abbreviations before using them, omit the
owner, and use consistent terms for equivalent access. For example:
`Finance reporting - Read reports`. Names containing the standalone terms
`Access Package`, `AP`, or `SG` (including underscore/hyphen-separated terms)
are rejected. Business meaning, owner names and consistent terminology require
human review; the script does not guess names from groups.

`-FallbackApproverIds` is mandatory and accepts one or more user IDs. These
users receive manager-stage requests when the recipient has no manager.
Explicit `AppOwnerIds` replaces, rather than supplements, the default owners.
For CSV, supplied owners create an optional second approval stage at any group
size. Without either, the CSV policy has only the manager stage.
`-DefaultAppOwnerIds` is not used to fill blank XLSX stages.

## Policy defaults

| Setting | Value |
| --- | --- |
| Eligible recipients | All members (excluding guests) (`allMemberUsers`). |
| Requesters | Self, authorized administrators, and the recipient's direct manager. |
| Request justification | Required. |
| Approval | Required for new assignments and extensions/updates. |
| CSV approval stages | Direct manager with fallback users, followed by application-owner users if supplied, regardless of group size. |
| XLSX approval stages | Exactly the populated approval-stage columns, regardless of group size. |
| Approval deadline | 14 days per stage; no escalation. |
| Approver justification | Not required; requester justification is required. |
| Assignment email notifications | Enabled. |
| Expiration | 365 days (`P365D`). |
| Request a custom timeline | Enabled. |
| Extend access | Enabled, with approval. |
| Access reviews | Disabled. |

Disabled reviews are sent as `reviewSettings: null`; an incomplete review
settings object can cause Graph to reject policy creation with
`ArgumentNullException`. Non-escalating approval stages explicitly include
empty escalation collections and a zero escalation duration.

The transitive **user** count includes users in nested groups, not nested group
objects, devices or service principals. The Graph advanced count query uses
eventual consistency. Group size does not determine approval stages. Manager approval
uses the recipient's manager (`targetManager`), not the manager of the person
submitting an on-behalf-of request.

The portal's "For users, service principals and agent identities in your
directory" heading has separate eligibility choices. This script selects
**All members (excluding guests)**, not all service principals or agents.
It does not enable requests from guests or users outside the directory.
Catalog visibility behavior is unchanged: new catalogs are externally visible,
but this does not make guests eligible under the member-only policy. Existing
catalog visibility is not changed. Existing packages and policies are skipped,
not migrated to the member-only scope.

### Required manual manager-bypass step

The public Graph v1.0 and beta **policy schemas do not expose** the portal
setting **Bypass approval stage if manager is the requester and approver**.
After creation, edit each new policy with a manager approval stage in the Entra
portal and enable this setting. Its created/previewed result includes `ManualAction`, and the script
prints a warning. A `Created` result means the Graph-configurable settings were
created, **not** that this manual step is complete. It never simulates bypass
by disabling manager approval for everyone.

## Results and failures

Each row emits an object with package/policy/catalog IDs, user count, stage
count, status, reason and the required manual action. Skipped/failed rows
and totals are printed separately without mixing summary text into the result
pipeline, so results can be captured or exported. Previewed and created
packages show the approvers for each stage, including the fallback if the
recipient has no manager. The results include that text in `ApprovalStageDetails`
alongside the numeric `ApprovalStages` count:

```powershell
$results = .\Create-AccessPackage.ps1 -CsvPath '.\access-packages.csv' `
    -CatalogName 'Business applications' `
    -FallbackApproverIds '55555555-5555-5555-5555-555555555555'
$results | Export-Csv '.\creation-results.csv' -NoTypeInformation
```

For failed Graph requests, `Reason` includes the HTTP method and URI, the
original exception, Graph error details (including its error code and message),
and the HTTP status and request ID when available. The response body is read
when the SDK does not populate `ErrorDetails`. Request bodies and authorization
headers are not logged. To inspect multiline errors without table truncation:

```powershell
$results | Where-Object Status -eq 'Failed' | Format-List Name, GroupId, Reason
```

Row failures do not stop subsequent rows. Authentication, missing columns,
missing fallback users and catalog ambiguity fail before provisioning. Catalogs
are created lazily after the first valid row is approved. Resource import
readiness is checked with a bounded retry before package creation.

If resource binding or policy creation fails, the script attempts to delete
**only the package it just created**. Shared catalogs and catalog resources
are retained. Rollback failures retain the package ID in the failed result and
explicitly request manual cleanup. A network failure during a create request
can leave server-side state without returning its ID; inspect the portal before
rerunning. Skipping an existing package does not validate or repair it.

## Tests

Offline Pester tests mock Graph and never connect to a tenant:

```powershell
Invoke-Pester -Path '.\tests\Create-AccessPackage.Tests.ps1' -Output Detailed
```

Graph references:
[assignment policies](https://learn.microsoft.com/en-us/graph/api/entitlementmanagement-post-assignmentpolicies?view=graph-rest-1.0),
[requester settings](https://learn.microsoft.com/en-us/graph/api/resources/accesspackageassignmentrequestorsettings?view=graph-rest-1.0),
[target manager](https://learn.microsoft.com/en-us/graph/api/resources/targetmanager?view=graph-rest-1.0),
and [catalog resource requests](https://learn.microsoft.com/en-us/graph/api/entitlementmanagement-post-resourcerequests?view=graph-rest-1.0).
