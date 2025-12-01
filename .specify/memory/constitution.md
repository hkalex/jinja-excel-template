
<!--
Sync Impact Report
Version change: N/A → 1.0.0
Modified principles: [PRINCIPLE_1_NAME] → Code Quality; [PRINCIPLE_2_NAME] → Testing Standards; [PRINCIPLE_3_NAME] → User Experience & API Stability; [PRINCIPLE_4_NAME] → Performance Requirements; [PRINCIPLE_5_NAME] → Simplicity & Maintainability
Added sections: Additional Constraints, Development Workflow
Removed sections: none
Templates updated: .specify/templates/plan-template.md (✅ updated), .specify/templates/tasks-template.md (✅ updated)
Follow-up TODOs: None
-->

# Jinja Excel Template Engine Constitution

## Core Principles

### I. Code Quality (NON-NEGOTIABLE)

Every contribution must maintain or improve the repository's code quality. The Code Quality principle requires that:

- Public APIs MUST be documented, typed (PEP 484), and include example usage in docstrings or the documentation site.
- Code changes MUST be reviewed and approved by at least one maintainer; high-risk or breaking changes REQUIRE two approvals.
- All code MUST pass static analysis (mypy), linters (flake8 or equivalent), and be formatted (Black + isort) prior to merging.
- PRs SHOULD be limited in size (target < 400 lines changed) to ease review and testing; larger changes MUST include a design summary and automated tests.

Rationale: enforcing code quality reduces maintenance burden, improves onboarding, and reduces regressions.

### II. Testing Standards (NON-NEGOTIABLE)

Tests are the project's most reliable safety net. This principle requires that:

- New features and bug fixes MUST include unit tests and, where applicable, integration or contract tests (pytest is the standard test runner).
- Tests MUST be written BEFORE the implementation for new behavior (Test-First / TDD) whenever possible. For legacy bug fixes, add a regression test first.
- The project recommends a coverage baseline of >= 95% for the repository; new code MUST not reduce this baseline without documented reasoning and approvals.
- All tests MUST run reliably on CI; flaky tests are unacceptable and must be fixed or quarantined with a clear plan and timeline.

Rationale: strong coverage and well-scoped tests reduce regressions and enable safe refactoring.

### III. User Experience & API Stability (MUST)

For a library, users rely on API stability and clear errors. This principle requires that:

- Public interface stability follows semantic versioning. Backwards-compatible changes SHOULD be minor or patch versions. MAJOR versions are required for breaking API changes.
- API ergonomics should be consistent and predictable: similar operations use consistent function/method names, parameter orders, and return semantics.
- Errors MUST be explicit and actionable: exceptions and messages should clearly describe the failure and where possible provide remediation guidance.
- Defaults MUST favor convenience and safety; new optional features are enabled opt-in by default to avoid breaking consumers.

Rationale: predictable and consistent APIs lower the integration cost for downstream projects.

### IV. Performance & Resource Budget (SHOULD)

The project must be efficient and predictable for typical workloads. The Performance principle requires:

- Performance goals MUST be established per feature (e.g., P95 latency for generating Excel from 1k–5k rows should be documented in the plan and validated by benchmarks).
- Each feature that impacts runtime or memory footprint MUST include a benchmark or micro-benchmark and a performance test. Performance regressions that exceed documented thresholds MUST be justified.
- For large data processing the implementation SHOULD stream data where possible to avoid excessive memory use and provide upper bounds for resource usage.

Rationale: measurable performance requirements let users plan capacity and ensure the library remains reliable at scale.

### V. Simplicity & Maintainability (MUST)

Favor simple and maintainable designs over premature optimization or unnecessary complexity.

- Designs SHOULD start with the simplest solution that meets user needs; avoid adding features without clear value.
- Favor readable code, small functions, and clear module boundaries to facilitate testing and review.
- Reuse helpers in `jinja_excel_template/helper/` rather than duplicating logic.

Rationale: simplicity reduces technical debt and enhances long-term maintainability.

## Additional Constraints

General project requirements that apply to all features:

- Supported Python: 3.11+
- Formatting: Black and isort enforced via pre-commit
- Static typing: mypy checks must pass for public APIs
- Linting: flake8 or a comparable linter enforced on CI
- CI gates: unit tests, integration tests, static checks, and performance checks must pass before merging
- Documentation: public functions and important behaviors must be documented; quickstarts and examples should exist for major features

Rationale: explicit tooling and constraints maintain consistency across the project and ease contributor onboarding.

## Development Workflow & Review

Contributors and maintainers must follow a predictable workflow:

- Branching: Use feature branches and a naming scheme like `feat/<brief>` or `fix/<brief>`.
- Pull Requests: Open PRs should be small and focused; large changes must include a design doc.
- Reviews: At least one maintainer review required to merge; two reviewers for MAJOR changes.
- Tests: Each PR must include tests (unit or integration) that validate the change; tests should be written before the implementation when feasible.
- CI: PRs must pass all checks: lints, formatting, type checks, tests, and performance benchmarks.

Rationale: consistent workflow reduces merge friction and improves code health.

## Governance

The Constitution governs how the project operates and how changes are adopted.

- Amendments: Any changes to the Constitution must be proposed via PR with a summary of the change, the reason, and a migration path if required. Amendments must be approved by two maintainers.
- Versioning: Follow semantic versioning for the constitution. MAJOR changes indicate breaking governance redefinitions; MINOR changes add principles and guidance; PATCHes are clarifications and non-functional wording updates.
- Compliance: The templates in `.specify/templates/` incorporate a Constitution Check to ensure new work adheres to these principles.

**Version**: 1.0.0 | **Ratified**: 2025-12-01 | **Last Amended**: 2025-12-01

<!--
Sync Impact Report
Version change: N/A → 1.0.0
Modified principles: [PRINCIPLE_1_NAME] → Code Quality; [PRINCIPLE_2_NAME] → Testing Standards; [PRINCIPLE_3_NAME] → User Experience & API Stability; [PRINCIPLE_4_NAME] → Performance Requirements; [PRINCIPLE_5_NAME] → Simplicity & Maintainability
Added sections: Additional Constraints, Development Workflow
Removed sections: none
Templates updated: .specify/templates/plan-template.md (✅ updated), .specify/templates/tasks-template.md (✅ updated)
Follow-up TODOs: None
-->

# Jinja Excel Template Engine Constitution

## Core Principles

### I. Code Quality (NON-NEGOTIABLE)

Every contribution must maintain or improve the repository's code quality. The Code Quality principle requires that:

- Public APIs MUST be documented, typed (PEP 484), and include example usage in docstrings or the documentation site.
- Code changes MUST be reviewed and approved by at least one maintainer; high-risk or breaking changes REQUIRE two approvals.
- All code MUST pass static analysis (mypy), linters (flake8 or equivalent), and be formatted (Black + isort) prior to merging.
- PRs SHOULD be limited in size (target < 400 lines changed) to ease review and testing; larger changes MUST include a design summary and automated tests.

Rationale: enforcing code quality reduces maintenance burden, improves onboarding, and reduces regressions.

### II. Testing Standards (NON-NEGOTIABLE)

Tests are the project's most reliable safety net. This principle requires that:

- New features and bug fixes MUST include unit tests and, where applicable, integration or contract tests (pytest is the standard test runner).
- Tests MUST be written BEFORE the implementation for new behavior (Test-First / TDD) whenever possible. For legacy bug fixes, add a regression test first.
- The project recommends a coverage baseline of >= 95% for the repository; new code MUST not reduce this baseline without documented reasoning and approvals.
- All tests MUST run reliably on CI; flaky tests are unacceptable and must be fixed or quarantined with a clear plan and timeline.

Rationale: strong coverage and well-scoped tests reduce regressions and enable safe refactoring.

### III. User Experience & API Stability (MUST)

For a library, users rely on API stability and clear errors. This principle requires that:

- Public interface stability follows semantic versioning. Backwards-compatible changes SHOULD be minor or patch versions. MAJOR versions are required for breaking API changes.
- API ergonomics should be consistent and predictable: similar operations use consistent function/method names, parameter orders, and return semantics.
- Errors MUST be explicit and actionable: exceptions and messages should clearly describe the failure and where possible provide remediation guidance.
- Defaults MUST favor convenience and safety; new optional features are enabled opt-in by default to avoid breaking consumers.

Rationale: predictable and consistent APIs lower the integration cost for downstream projects.

### IV. Performance & Resource Budget (SHOULD)

The project must be efficient and predictable for typical workloads. The Performance principle requires:

- Performance goals MUST be established per feature (e.g., P95 latency for generating Excel from 1k–5k rows should be documented in the plan and validated by benchmarks).
- Each feature that impacts runtime or memory footprint MUST include a benchmark or micro-benchmark and a performance test. Performance regressions that exceed documented thresholds MUST be justified.
- For large data processing the implementation SHOULD stream data where possible to avoid excessive memory use and provide upper bounds for resource usage.

Rationale: measurable performance requirements let users plan capacity and ensure the library remains reliable at scale.

### V. Simplicity & Maintainability (MUST)

Favor simple and maintainable designs over premature optimization or unnecessary complexity.

- Designs SHOULD start with the simplest solution that meets user needs; avoid adding features without clear value.
- Favor readable code, small functions, and clear module boundaries to facilitate testing and review.
- Reuse helpers in `jinja_excel_template/helper/` rather than duplicating logic.

Rationale: simplicity reduces technical debt and enhances long-term maintainability.

## Additional Constraints

General project requirements that apply to all features:

- Supported Python: 3.11+
- Formatting: Black and isort enforced via pre-commit
- Static typing: mypy checks must pass for public APIs
- Linting: flake8 or a comparable linter enforced on CI
- CI gates: unit tests, integration tests, static checks, and performance checks must pass before merging
- Documentation: public functions and important behaviors must be documented; quickstarts and examples should exist for major features

Rationale: explicit tooling and constraints maintain consistency across the project and ease contributor onboarding.

## Development Workflow & Review

Contributors and maintainers must follow a predictable workflow:

- Branching: Use feature branches and a naming scheme like `feat/<brief>` or `fix/<brief>`.
- Pull Requests: Open PRs should be small and focused; large changes must include a design doc.
- Reviews: At least one maintainer review required to merge; two reviewers for MAJOR changes.
- Tests: Each PR must include tests (unit or integration) that validate the change; tests should be written before the implementation when feasible.
- CI: PRs must pass all checks: lints, formatting, type checks, tests, and performance benchmarks.

Rationale: consistent workflow reduces merge friction and improves code health.

## Governance

The Constitution governs how the project operates and how changes are adopted.

- Amendments: Any changes to the Constitution must be proposed via PR with a summary of the change, the reason, and a migration path if required. Amendments must be approved by two maintainers.
- Versioning: Follow semantic versioning for the constitution. MAJOR changes indicate breaking governance redefinitions; MINOR changes add principles and guidance; PATCHes are clarifications and non-functional wording updates.
- Compliance: The templates in `.specify/templates/` incorporate a Constitution Check to ensure new work adheres to these principles.

**Version**: 1.0.0 | **Ratified**: 2025-12-01 | **Last Amended**: 2025-12-01
# [PROJECT_NAME] Constitution
<!-- Example: Spec Constitution, TaskFlow Constitution, etc. -->

## Core Principles

### [PRINCIPLE_1_NAME]
<!-- Example: I. Library-First -->
[PRINCIPLE_1_DESCRIPTION]
<!-- Example: Every feature starts as a standalone library; Libraries must be self-contained, independently testable, documented; Clear purpose required - no organizational-only libraries -->

### [PRINCIPLE_2_NAME]
<!-- Example: II. CLI Interface -->
[PRINCIPLE_2_DESCRIPTION]
<!-- Example: Every library exposes functionality via CLI; Text in/out protocol: stdin/args → stdout, errors → stderr; Support JSON + human-readable formats -->

### [PRINCIPLE_3_NAME]
<!-- Example: III. Test-First (NON-NEGOTIABLE) -->
[PRINCIPLE_3_DESCRIPTION]
<!-- Example: TDD mandatory: Tests written → User approved → Tests fail → Then implement; Red-Green-Refactor cycle strictly enforced -->

### [PRINCIPLE_4_NAME]
<!-- Example: IV. Integration Testing -->
[PRINCIPLE_4_DESCRIPTION]
<!-- Example: Focus areas requiring integration tests: New library contract tests, Contract changes, Inter-service communication, Shared schemas -->

### [PRINCIPLE_5_NAME]
<!-- Example: V. Observability, VI. Versioning & Breaking Changes, VII. Simplicity -->
[PRINCIPLE_5_DESCRIPTION]
<!-- Example: Text I/O ensures debuggability; Structured logging required; Or: MAJOR.MINOR.BUILD format; Or: Start simple, YAGNI principles -->

## [SECTION_2_NAME]
<!-- Example: Additional Constraints, Security Requirements, Performance Standards, etc. -->

[SECTION_2_CONTENT]
<!-- Example: Technology stack requirements, compliance standards, deployment policies, etc. -->

## [SECTION_3_NAME]
<!-- Example: Development Workflow, Review Process, Quality Gates, etc. -->

[SECTION_3_CONTENT]
<!-- Example: Code review requirements, testing gates, deployment approval process, etc. -->

## Governance
<!-- Example: Constitution supersedes all other practices; Amendments require documentation, approval, migration plan -->

[GOVERNANCE_RULES]
<!-- Example: All PRs/reviews must verify compliance; Complexity must be justified; Use [GUIDANCE_FILE] for runtime development guidance -->

**Version**: [CONSTITUTION_VERSION] | **Ratified**: [RATIFICATION_DATE] | **Last Amended**: [LAST_AMENDED_DATE]
<!-- Example: Version: 2.1.1 | Ratified: 2025-06-13 | Last Amended: 2025-07-16 -->
