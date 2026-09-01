# Installation
## Local Installation
### Prod installation
```bash
uv sync
```
### Dev installation
```bash
uv sync --dev
```
### Check for Updates
```
uv lock --upgrade-package espy && uv sync
```
## Adding Dependencies

!!! warning 

    When adding dependencies, please evaluate:

    - [ ] Is this package well-maintained?
    - [ ] Is this package necessary?
    - [ ] Should this package be added as a production dependency, or a dev dependency?

### Adding a prod dependency
```bash
uv add <dependency name>
```

### Adding a dev dependency
```bash
uv add <dev dependency name>
```