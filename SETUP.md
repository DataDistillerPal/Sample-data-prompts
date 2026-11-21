# Power BI Analytics Agent Kit: Setup From Scratch

A VS Code + GitHub Copilot workspace with a team of Power BI agents, a data-privacy policy, a knowledge base of
your semantic models, and a skills library you keep growing.

---

## What's in the kit

```
powerbi-agent-kit/
├── .github/
│   ├── copilot-instructions.md            ← core rules: ask-before-assuming, safety, validation
│   ├── instructions/
│   │   └── data-policy.instructions.md    ← PII & data boundaries (applies to everything)
│   ├── agents/
│   │   ├── analytics-engineer.agent.md    ← MAIN agent: data engineering + analysis (PQ, SQL, Python, modelling)
│   │   ├── dax-expert.agent.md            ← DAX, minimal CALCULATE, standard date table
│   │   ├── model-reviewer.agent.md        ← read-only review (best practice + privacy)
│   │   ├── model-documenter.agent.md      ← descriptions, format strings, data dictionary
│   │   └── catalog-builder.agent.md       ← builds the knowledge base of your models
│   ├── skills/                            ← GROWING LIBRARY. Add skills here.
│   │   ├── README.md                      ← how to add a skill + template
│   │   ├── dax-best-practices/SKILL.md
│   │   ├── power-query-best-practices/SKILL.md
│   │   ├── sql-best-practices/SKILL.md
│   │   ├── python-analytics/SKILL.md
│   │   └── semantic-model-design/SKILL.md
│   └── prompts/
│       ├── add-skill.prompt.md            ← /add-skill
│       ├── onboard-model.prompt.md        ← /onboard-model
│       └── review-model.prompt.md         ← /review-model
├── knowledge/                             ← the agents' memory of YOUR environment (metadata only)
│   ├── README.md
│   ├── standards/date-table.md            ← FILL IN: your standard date table
│   ├── glossary.md                        ← FILL IN: KPI definitions
│   ├── model-catalog.md                   ← index of models
│   ├── models/_TEMPLATE.md                ← one file per model (written by catalog-builder)
│   ├── pii-register.md                    ← PII columns across models
│   └── learned-rules.md                   ← corrections you've taught the agents
├── models/                                ← (create) put PBIP projects here
├── .vscode/settings.json                  ← skills on, instruction files on, no auto-approve
├── .vscode/mcp.json                       ← optional read-only Power BI server
└── .gitignore                             ← keeps data files & secrets out of Git
```

### How the pieces fit
- **Instructions + policy** apply to every chat, every time.
- **Agents** are personas you pick from the agent dropdown. Each has handoff buttons to the next agent.
- **Skills** load automatically when your request matches (DAX question → DAX skill).
- **Knowledge** tells the agents what *your* models, date table and KPIs look like.
- **Prompts** are `/shortcuts` for repeated workflows.

**Agent flow:** `catalog-builder` (learn the model) → `model-reviewer` (find issues) → `dax-expert` / `analytics-engineer` (fix) → `model-documenter` (document) → back to `model-reviewer`.

---

## Phase 0: Checks with IT (do first)

- [ ] **GitHub Copilot Business or Enterprise** seat on your work GitHub account
- [ ] Org policy **"MCP servers in Copilot"** enabled (GitHub org/enterprise → Copilot → Policies)
- [ ] **Agent skills / preview features** allowed by your Copilot policy (ask if unsure)
- [ ] Allowed to use AI tools on client models? Which clients? (Get it in writing.)
- [ ] Power BI **Pro / PPU** license; for Fabric models: **Contributor or Build** permission on a **dev** workspace
- [ ] A **dev copy** of at least one model to practise on
- [ ] (Optional) **Git** repo location approved (Azure DevOps / GitHub Enterprise) for PBIP projects
- [ ] (Optional) Python allowed locally? Or use Fabric notebooks?

---

## Phase 1: Install the tools (about 20 minutes)

1. **VS Code**: https://code.visualstudio.com (keep it updated; agents, skills and handoffs need a recent version).
2. **Git for Windows**: https://git-scm.com (needed for version control of PBIP files).
3. **Power BI Desktop**: latest version.
   - File → Options → Preview features → enable **Power BI Project (.pbip) save option** and the
     **TMDL** options (if still listed as preview in your version).
4. In VS Code → Extensions (`Ctrl+Shift+X`) install:
   - **GitHub Copilot** (includes Copilot Chat)
   - **Power BI Modeling MCP** (publisher: Microsoft)
   - *(Optional)* **Python** (Microsoft), **SQL Server (mssql)**, **Power Query / M Language**
5. Sign in to GitHub in VS Code (Accounts icon, bottom-left) with your **work** account.
6. Open Copilot Chat (`Ctrl+Alt+I`), switch mode to **Agent**, click the 🔧 tools icon, and confirm
   **powerbi-modeling-mcp** is listed. Accept its EULA when prompted on first use.

---

## Phase 2: Create your workspace (10 minutes)

1. Create a folder, e.g. `C:\Work\powerbi-agent-workspace\`.
2. Unzip the kit **contents** into it, so `.github`, `knowledge`, `.vscode`, `.gitignore`, `SETUP.md` sit at the root.
   (Hidden folders: Windows Explorer → View → Show → **Hidden items**.)
3. Create a `models\` subfolder.
4. Open a terminal in the folder and initialise Git:
   ```
   git init
   git add .
   git commit -m "Initial agent kit"
   ```
5. In VS Code: **File → Open Folder…** → select the workspace folder.
   ⚠️ Always open *this* folder. Agents only load when the folder containing `.github` is the open folder.

---

## Phase 3: Teach it your environment (30–60 minutes, once)

1. **`knowledge/standards/date-table.md`**: fill in your standard date table: table name, date column, fiscal year,
   columns. *This is what stops agents creating calendar tables.*
2. **`knowledge/glossary.md`**: add your top 10 KPI definitions.
3. **`.github/copilot-instructions.md`** §4: fill in your sources, SQL dialect, Python environment, prod workspace names.
4. **`.github/instructions/data-policy.instructions.md`**: check the PII list and the small-group threshold `[10]`.
5. **Save your models as PBIP** into `models\` (Power BI Desktop → File → Save as → *Power BI project*). Commit.
6. Commit: `git add . && git commit -m "Environment knowledge"`

---

## Phase 4: Verify everything loaded (5 minutes)

| Check | How | Expected |
|---|---|---|
| Agents | Copilot Chat → agent dropdown | analytics-engineer, dax-expert, model-reviewer, model-documenter, catalog-builder |
| Prompts | Type `/` in chat | add-skill, onboard-model, review-model |
| Skills | Type `/skills` in chat (or Command Palette → *Chat: Open Customizations*) | the 5 skills listed |
| Instructions | Ask anything → expand **References** under the reply | `copilot-instructions.md` and `data-policy.instructions.md` |
| Power BI tools | Pick an agent → 🔧 tools | powerbi-modeling-mcp tools ticked |
| Policy works | Ask: *"Show me customer emails from Dim Customer"* | Agent stops and asks for `APPROVE PII` |
| Ask-first works | Ask: *"Create a sales YoY measure"* | Agent asks which date, fiscal/calendar, definition of sales, before writing |

If an agent says it has no Power BI tools: open its `.agent.md` → click **Configure Tools** (above the frontmatter)
→ tick the Power BI server → save. VS Code writes the exact tool name.

---

## Phase 5: Build the knowledge base (ongoing)

1. Open your first model in Power BI Desktop (or know its Fabric workspace name).
2. Run **`/onboard-model`** → answer the agent's questions → review the proposed files → approve.
3. Open `knowledge/pii-register.md` → change `Suggested` to `Confirmed` or `Not PII`.
4. Repeat for your 3–5 most-used models. Commit after each.
5. Re-run `/onboard-model` after major model changes (it shows only the differences).

You don't need a report that "points to all reports". The catalog + one file per model (+ PBIP files in `models\`)
*is* how the agents know your models.

---

## Phase 6: First real tasks

1. **Review:** `/review-model` → read findings → use the handoff buttons.
2. **DAX:** pick **dax-expert** → *"I need Sales YoY % for the Contoso model."*
   Expect: clarifying questions → proposal using your date table → CALCULATE check → approval → validation query.
3. **Engineering:** pick **analytics-engineer** → *"The Sales table refresh is slow. Find out why."*
   Expect: it checks folding, source queries, and model size, then proposes fixes in the right layer (SQL / PQ / DAX).
4. **Analysis:** **analytics-engineer** → *"Which product categories drove the change in margin this quarter?"*
   Expect: clarifying questions → DAX/SQL aggregates → headline-first answer.

---

## Phase 7: Daily workflow

1. Open the workspace folder in VS Code; open the model in Power BI Desktop.
2. Pick the agent for the job (or start with **analytics-engineer**, which hands off as needed).
3. Answer its questions; approve its plan; approve each change.
4. Let it validate; save in Power BI Desktop.
5. `git commit` after each meaningful change. That's your undo button.
6. When you correct the agent, say yes to *"add this to learned-rules.md?"* if it's a standing rule.

---

## Phase 8: Grow the skills library

- **Quick:** `/add-skill <topic>` → answer questions → approve.
- **Manual:** see `.github/skills/README.md` (template included).
- **When to add one:** you've corrected the agent the same way twice; a new client has specific KPI rules;
  you've written a pattern you'll reuse (currency conversion, RLS, paginated reports).
- **Where to grab more:** the Awesome GitHub Copilot collection (awesome-copilot.github.com) and Microsoft's
  `powerbi-authoring` plugin (includes a semantic-model-authoring skill). Review any downloaded skill before using it.

---

## How the boundaries are enforced (be clear-eyed)

The policy file is **guidance the AI follows**, not a technical lock. Layer real controls on top:

| Layer | What it does | Set up |
|---|---|---|
| Data policy + instructions | Agent asks for `APPROVE PII`, masks output, limits rows | ✅ In kit |
| Tool approvals | You confirm every tool call (`chat.tools.autoApprove: false`) | ✅ In kit |
| Modeling MCP write confirmation | Server asks before each model edit | ✅ Default |
| Read-only server | `--readonly` blocks all writes | Optional: `.vscode/mcp.json` |
| Least-privilege account | Agent can only do what *your* account can do | Use dev workspaces; avoid admin accounts |
| OLS / RLS in dev models | Hide PII columns from the account you test with | Ask model owner |
| Git | Every change is reviewable and reversible | ✅ Phase 2 |
| `.gitignore` | Data files never committed | ✅ In kit |
| Hooks (advanced, preview) | Scripts that can block specific tool calls automatically | Later: `.github/hooks/*.json` |

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Agents/prompts missing | Open the folder that **contains** `.github`; `Ctrl+Shift+P` → *Developer: Reload Window* |
| Skills not listed | Update VS Code; check `chat.useAgentSkills` is on; check your org allows skills; folder name must match `name:` |
| Power BI tools missing | Org policy "MCP servers in Copilot" off, or extension disabled |
| Agent has no Power BI tools | `.agent.md` → **Configure Tools** → tick the server |
| Can't connect to Desktop | File must be open in Power BI Desktop; use the exact file name |
| Agent ignores policy | Check References shows `data-policy.instructions.md`; check `useInstructionFiles` setting; report it, and tighten the account permissions |
| Agent invents a date table | Fill in `knowledge/standards/date-table.md`; add a line to `learned-rules.md` |
| Answers are shallow | Pick a stronger reasoning model in the chat model picker; make sure the model's knowledge file exists |
