# Access package creation

[Create-AccessPackage.ps1](./Create-AccessPackage.ps1) creates one group-based
access package and request policy per CSV row. It reuses the mandatory
`-CatalogName` or creates a published catalog if no exact match exists.
Ambiguous catalog names stop the run; unpublished catalogs must be published
first. Existing packages with the same name in that catalog are skipped,
never overwritten. Duplicate names within a CSV are processed once.

## Prerequisites and use

Connect to Microsoft Graph before running the script. Requires the
`Microsoft.Graph.Authentication` module, appropriate Entra ID Governance
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

## CSV and naming

Start with [access-packages.example.csv](./access-packages.example.csv).
The comma-delimited CSV requires all four headers:

| Column | Meaning |
| --- | --- |
| `Name` | Business-friendly application/service name and the access provided. |
| `Description` | Explain what the requester can do with this access. |
| `GroupId` | One Entra group object ID, not its name. |
| `AppOwnerIds` | Explicit approver **user** object IDs separated by semicolons. Blank uses `-DefaultAppOwnerIds`. |

Use the application/service name requesters recognize, describe the access
instead of the Entra group, explain abbreviations before using them, omit the
owner, and use consistent terms for equivalent access. For example:
`Finance reporting - Read reports`. Names containing the standalone terms
`Access Package`, `AP`, or `SG` (including underscore/hyphen-separated terms)
are rejected. Business meaning, owner names and consistent terminology require
human review; the script does not guess names from groups.

`-FallbackApproverIds` is mandatory and accepts one or more user IDs. These
users receive manager-stage requests when the recipient has no manager.
Explicit `AppOwnerIds` replaces, rather
than supplements, the default owners. Owners are validated when supplied;
at least one is required for a group with more than 500 users.

## Policy defaults

| Setting | Value |
| --- | --- |
| Eligible recipients | All members (excluding guests) (`allMemberUsers`). |
| Requesters | Self, authorized administrators, and the recipient's direct manager. |
| Request justification | Required. |
| Approval | Required for new assignments and extensions/updates. |
| Up to and including 500 transitive users | Direct manager, with explicit fallback users. |
| More than 500 transitive users | Direct manager, then the explicit application-owner users. |
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
eventual consistency. Stage selection is a creation-time snapshot; the script
does not continuously adjust policies as membership changes. Manager approval
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
After creation, edit each new policy in the Entra portal and enable this
setting. Every created/previewed result includes `ManualAction`, and the script
prints a warning. A `Created` result means the Graph-configurable settings were
created, **not** that this manual step is complete. It never simulates bypass
by disabling manager approval for everyone.

## Results and failures

Each row emits an object with package/policy/catalog IDs, user count, stage
count, status, reason and the required manual action. Skipped/failed rows
and totals are printed separately without mixing summary text into the result
pipeline, so results can be captured or exported:

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
