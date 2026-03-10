diff --git a/README.md b/README.md
index 0d7871fa4f65a84018fe86348807951fb2c2f7e5..a941ecdc91494706d681bc385c77d9514832c073 100644
--- a/README.md
+++ b/README.md
@@ -1 +1,88 @@
-# projectscheduler
\ No newline at end of file
+# Nequette Production Scheduler (SPFx)
+
+This repository now includes a starter SPFx scheduler implementation with core scaffolding for:
+
+- scheduler web part rendering
+- scheduler engine data normalization + timeline rows
+- SharePoint list CRUD service abstraction
+- provisioning + JSON import PowerShell scripts
+- dedicated build instructions
+
+## Included deliverables
+
+- Web part entry point: `src/webparts/productionScheduler/ProductionSchedulerWebPart.ts`
+- Scheduler engine port scaffold: `src/webparts/productionScheduler/SchedulerEngine.ts`
+- SharePoint data service scaffold: `src/webparts/productionScheduler/SharePointDataService.ts`
+- Stylesheet: `src/webparts/productionScheduler/SchedulerStyles.scss`
+- Provisioning script: `scripts/Create-SchedulerLists.ps1`
+- Migration/import script: `scripts/import-scheduler-json.ps1`
+- Build guide: `BUILD-INSTRUCTIONS.md`
+
+## Recommended development environment
+
+- Node.js **18 LTS**
+- npm (bundled with Node.js)
+- Global tooling:
+
+```bash
+npm install -g yo gulp @microsoft/generator-sharepoint
+```
+
+## Build and run workflow
+
+1. Install dependencies:
+
+   ```bash
+   npm install
+   ```
+
+2. Run local development:
+
+   ```bash
+   gulp serve
+   ```
+
+   Then open:
+
+   ```text
+   https://yourtenant.sharepoint.com/_layouts/15/workbench.aspx
+   ```
+
+3. Build production bundle:
+
+   ```bash
+   gulp bundle --ship
+   ```
+
+4. Package solution:
+
+   ```bash
+   gulp package-solution --ship
+   ```
+
+   Output package:
+
+   ```text
+   sharepoint/solution/nequette-production-scheduler.sppkg
+   ```
+
+## Deployment
+
+1. SharePoint Admin Center
+2. App Catalog → Apps for SharePoint
+3. Upload `.sppkg`
+4. Add **Production Scheduler** web part to a page
+
+## Script usage
+
+### Provision scheduler lists
+
+```powershell
+./scripts/Create-SchedulerLists.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/operations"
+```
+
+### Import scheduler JSON data
+
+```powershell
+./scripts/import-scheduler-json.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/operations" -JsonPath "./data/scheduler.json"
+```