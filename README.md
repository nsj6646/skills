# skills

Personal repository for creating and managing agent skills.

## Structure

```text
.
├── AGENTS.md              # Repository and agent skill conventions
├── docs/agents/           # Issue tracker, triage labels, and domain doc rules
├── scripts/
│   └── link-skills.sh     # Symlink skills/* into the local agent skills directory
└── skills/                # Skill packages (<skill-name>/SKILL.md)
```

## Installation & Local Sync

### Local development (symlink)

Link all skills in `skills/` into `~/.gemini/antigravity-cli/skills` (override target with `SKILLS_TARGET_DIR` or first argument):

```bash
./scripts/link-skills.sh
```

### Install via Skills CLI

```bash
npx skills add nsj6646/skills -g
```

## Skills

| Skill | Description |
| ----- | ----------- |
| _None yet_ | Add skills under `skills/<skill-name>/SKILL.md`. |
