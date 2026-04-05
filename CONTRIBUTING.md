# Contributing

## Prerequisites

- Node.js 20+
- Rust toolchain installed through `rustup`

## Local development

```bash
npm install
npm run build
npm test
```

## Quality checks

```bash
npm run format:check
npm run test:rust
npm run test:ts
```

## Notes

- Keep public package code, logs, and technical errors in English.
- Do not commit `target/` or local `.node` artifacts.
- Update `CHANGELOG.md` for user-visible changes.
