# Copilot Instructions

- Always use Conventional Commits for every commit message.
- Use this format: `<type>(<optional-scope>): <description>`.
- Allowed types: `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`, `revert`.
- Keep the description short and imperative (for example: `fix(api): handle null user id`).
- Mark breaking changes with `!` after type/scope (for example: `feat(auth)!: remove legacy token flow`) and include `BREAKING CHANGE:` in the commit body when applicable.

## Changelog

- **Always update `CHANGELOG.md`** when making any user-facing change (feat, fix, perf, refactor, or breaking change).
- Add entries under the `[Unreleased]` section at the top of the file.
- Group entries by type: `Added`, `Changed`, `Fixed`, `Removed`, `Security`, `Deprecated`.
- Reference the related PR number as a link (e.g. `([#123])`).
- Do **not** add changelog entries for pure dependency bumps handled by Dependabot — those are batched by a human when relevant.
