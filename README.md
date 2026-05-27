# Claude Code GSD Agents

Collection of specialized sub-agents for the **GSD (Get Shit Done)** workflow system, used with [Claude Code](https://claude.ai/code).

These agent definition files extend Claude Code with structured, role-specific behaviors for software development workflows — planning, execution, review, debugging, and more.

---

## What are GSD Agents?

Each `.md` file defines a sub-agent with:
- A specific **role** and **responsibilities**
- A set of allowed **tools**
- Structured **prompts** that guide the agent's behavior

They are spawned automatically by GSD slash commands (e.g. `/gsd:plan-phase`, `/gsd:execute-phase`, `/gsd:debug`) and run in isolated contexts to keep the main conversation clean.

---

## Agents

### Planning & Execution
| Agent | Description |
|-------|-------------|
| `gsd-planner` | Creates executable phase plans with task breakdown and dependency analysis |
| `gsd-executor` | Executes PLAN.md files atomically with per-task commits and state management |
| `gsd-plan-checker` | Verifies plans will achieve the phase goal before execution |
| `gsd-phase-researcher` | Researches how to implement a phase before planning |
| `gsd-pattern-mapper` | Maps new files to closest analogs in existing codebase |

### Review & Quality
| Agent | Description |
|-------|-------------|
| `gsd-code-reviewer` | Reviews source files for bugs, security issues, and code quality |
| `gsd-code-fixer` | Applies fixes to findings from REVIEW.md |
| `gsd-verifier` | Validates that built features deliver what the phase promised |
| `gsd-integration-checker` | Checks cross-phase integration and E2E flows |
| `gsd-security-auditor` | Verifies threat mitigations exist in implemented code |
| `gsd-nyquist-auditor` | Fills validation gaps by generating tests and verifying coverage |

### Debugging
| Agent | Description |
|-------|-------------|
| `gsd-debugger` | Investigates bugs using the scientific method with checkpoint support |
| `gsd-debug-session-manager` | Manages multi-cycle debug checkpoint and continuation loops |

### Documentation
| Agent | Description |
|-------|-------------|
| `gsd-doc-writer` | Writes and updates project documentation |
| `gsd-doc-classifier` | Classifies planning documents (ADR, PRD, SPEC, DOC) |
| `gsd-doc-synthesizer` | Synthesizes classified planning docs into a consolidated context |
| `gsd-doc-verifier` | Verifies factual claims in generated docs against the live codebase |

### Research & Analysis
| Agent | Description |
|-------|-------------|
| `gsd-project-researcher` | Researches the domain ecosystem before roadmap creation |
| `gsd-domain-researcher` | Surfaces domain expert evaluation criteria and failure modes |
| `gsd-research-synthesizer` | Synthesizes outputs from parallel researcher agents |
| `gsd-assumptions-analyzer` | Deeply analyzes codebase assumptions with evidence |
| `gsd-codebase-mapper` | Explores codebase and writes structured analysis documents |
| `gsd-intel-updater` | Analyzes codebase and writes structured intel files |

### AI Integration
| Agent | Description |
|-------|-------------|
| `gsd-ai-researcher` | Researches AI framework docs for implementation-ready guidance |
| `gsd-framework-selector` | Interactive decision matrix for selecting the right AI/LLM framework |
| `gsd-eval-planner` | Designs structured evaluation strategies for AI phases |
| `gsd-eval-auditor` | Retroactive audit of AI phase evaluation coverage |

### UI & Frontend
| Agent | Description |
|-------|-------------|
| `gsd-ui-researcher` | Produces UI-SPEC.md design contracts for frontend phases |
| `gsd-ui-checker` | Validates UI-SPEC.md design contracts across 6 quality dimensions |
| `gsd-ui-auditor` | Retroactive 6-pillar visual audit of implemented frontend code |

### Project Management
| Agent | Description |
|-------|-------------|
| `gsd-roadmapper` | Creates project roadmaps with phase breakdown and requirement mapping |
| `gsd-advisor-researcher` | Researches gray area decisions and returns structured comparison tables |
| `gsd-user-profiler` | Analyzes session messages to produce a scored developer profile |

---

## Installation

Place the `.md` files in your Claude Code agents directory:

```
~/.claude/agents/
```

Claude Code will automatically discover and load them.

---

## Usage

These agents are invoked via GSD slash commands in Claude Code:

```
/gsd:plan-phase     → spawns gsd-planner, gsd-plan-checker
/gsd:execute-phase  → spawns gsd-executor
/gsd:code-review    → spawns gsd-code-reviewer
/gsd:debug          → spawns gsd-debug-session-manager, gsd-debugger
/gsd:verify-work    → spawns gsd-verifier
```

---

## Author

Antonio Teodoro Frutuoso — Thomson Reuters
