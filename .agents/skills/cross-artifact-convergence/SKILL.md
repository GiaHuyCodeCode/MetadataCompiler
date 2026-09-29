---
name: cross-artifact-convergence
description: Performs non-destructive cross-artifact consistency analysis across spec.md, plan.md, tasks.md, and codebase. Audits unbuilt work and appends remediation tasks to ensure 100% convergence.
---

# Cross-Artifact Convergence & Semantic Analysis

## Overview
Verifies consistency across the three core artifacts (`spec.md`, `plan.md`, `tasks.md`) before coding, and audits the actual codebase against the spec upon completion to generate remediation tasks for any unbuilt work.

## The 3-Way Semantic Alignment Model
Check that every single requirement key (`FR-###`, `US-###`, `AC-###`):
- Has an architectural component in `architecture.md` / `plan.md`.
- Has at least one `[TEST]` task in `tasks.md`.
- Has at least one `[IMPL]` task in `tasks.md`.

## Pre-Implementation Consistency Audit Workflow
1. Load `PRD.md`, `architecture.md`, `tasks.md`.
2. Check Entity Alignment, Layer Coverage (no dropped layers), and Edge Cases.
3. If issues found -> STOP and adjust tasks before coding.

## Post-Implementation Codebase Convergence Workflow
1. Verify actual code files exist with real logic (no leftover `// TODO`).
2. Run "Tests-Exist" check: every logic file has paired test file.
3. Verify test coverage meets or exceeds **80%** with HTML report.
4. If any AC is missing: automatically append **Remediation Tasks** to `tasks.md`.
