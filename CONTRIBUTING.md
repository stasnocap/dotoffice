# Contributing

## Commit messages

This repo uses [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>(<scope>): <summary>
```

- **Types:** `feat`, `fix`, `perf`, `refactor`, `test`, `docs`, `ci`, `build`, `chore`
- **Scopes:** `core`, `fonts`, `pdf`, `layout`, `docx`, `xlsx`, `pptx`, `converter`, `samples`
- Breaking changes: `feat(pdf)!: ...` or a `BREAKING CHANGE:` footer
- Reference issues in the footer, e.g. `Closes #2`

Examples:

```
feat(fonts): parse kerning from GPOS
fix(pdf): escape parentheses in document title
ci: validate sample PDF with qpdf
```

## Guidelines

- No native dependencies; everything must be managed .NET.
- No dependencies with GPL/AGPL or other non-permissive licenses.
- Test documents must be public domain or self-generated.
