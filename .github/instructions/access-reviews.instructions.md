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
Keep logs/reports private and reporting offline. Treat reported configuration as
computed/submitted metadata, not proof of live review state or remediation.
