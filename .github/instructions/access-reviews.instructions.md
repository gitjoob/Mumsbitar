---
applyTo: "Entra/IdentityGovernance/AccessReviews/**"
---

# AccessReviews

Follow the tool-local [AGENTS.md](../../Entra/IdentityGovernance/AccessReviews/AGENTS.md)
and use its [README.md](../../Entra/IdentityGovernance/AccessReviews/README.md)
as the policy and usage reference.

Keep changes confined to the requested behavior and preserve existing switches,
duplicate protection, nested-group coverage, reviewer routing and schedule rules.
Validate with Graph-mocked Pester tests; never use live tenant writes as tests.
Keep logs/reports private and generated HTML offline. Creation/offline reports
show computed/submitted metadata, not live review state. The separate status
collector uses GET only; distinguish current configuration/owners from recorded
reviewer evidence and runtime stages. Neither report proves remediation.
