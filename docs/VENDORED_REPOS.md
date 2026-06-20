# Vendored Source Repositories

The root repository vendors the source files from the cloned upstream projects
below as ordinary directories, not as Git submodules. This keeps a fresh clone
of this project self-contained.

## `docformat-gui`

- Upstream: `https://github.com/KaguraNanaga/docformat-gui`
- Branch at import: `main`
- Commit at import: `a5c4a341f1a3fe791bf6e16d281f3c6fa3729f72`
- Commit subject: `Speed up macOS notarization test builds`

## `HeuThesis_Overleaf`

- Upstream: `https://github.com/phhandong/HeuThesis_Overleaf`
- Branch at import: `master`
- Commit at import: `aad7bdd8e438490bcb5ad733782b9a0a3eff35d0`
- Commit subject: `add: pdf preview&concat info`

The original nested `.git` metadata directories were moved to a local backup
outside the root repository before the first root commit, so the vendored files
can be tracked as normal source files.

