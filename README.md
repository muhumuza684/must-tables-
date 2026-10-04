# MUST Tables

A Power BI custom visual for analytical tables: search, filter, sort, group, pivot, calculate, format and export.

Built by **Delight BI** - *Data Guided Decisions* - a Bryt Ma Tech UG business.

## Get the visual

- **Latest build:** open the **Actions** tab, pick the newest green run, and download the `must-tables-pbiviz` artifact (kept for 30 days).
- **Permanent releases:** see **Releases**. A release is created whenever a version tag such as `v1.0.0.50` is pushed.

Import it in Power BI Desktop: *Visualizations* > *...* > *Import a visual from a file*.

## Develop

```
npm ci            # install exactly what package-lock.json says
npm test          # unit tests
npm run ci        # the same gate GitHub Actions runs: type-check, tests, lint, package
npm run package   # builds dist/*.pbiviz
```

## Licence

Copyright (c) Bryt Ma Tech UG. All rights reserved.
