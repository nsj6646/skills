## Repository conventions

### Skill layout

- Place each skill in `skills/<skill-name>/SKILL.md`, using a lowercase hyphenated directory name matching the frontmatter `name`.
- Keep `SKILL.md` concise; move bulky reference docs to `skills/<skill-name>/references/` and helper scripts to `skills/<skill-name>/scripts/`.

### Authoring & maintenance

- Follow the `writing-for-agents` skill when creating or editing skills (keep `description` sharp for model-invoked skills, or set `disable-model-invocation: true` for manual-only skills).
- Update the skills table in `README.md` whenever a skill is added, renamed, or removed.
- Run `./scripts/link-skills.sh` after adding or renaming a skill to update local symlinks.

## Agent skills

### Issue tracker

Issues and specs are tracked in GitHub Issues (`nsj6646/skills`) via the `gh` CLI. See `docs/agents/issue-tracker.md`.

### Triage labels

Default triage labels (`needs-triage`, `needs-info`, `ready-for-agent`, `ready-for-human`, `wontfix`). See `docs/agents/triage-labels.md`.

### Domain docs

Single-context layout (`GLOSSARY.md` and `docs/adr/` at the repo root). See `docs/agents/domain.md`.
