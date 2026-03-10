# Build Instructions

## Prerequisites

- Node.js 18 LTS
- npm
- SPFx tooling:

```bash
npm install -g yo gulp @microsoft/generator-sharepoint
```

## Install dependencies

```bash
npm install
```

## Run local development workbench

```bash
gulp serve
```

Open:

```text
https://yourtenant.sharepoint.com/_layouts/15/workbench.aspx
```

## Build production assets

```bash
gulp bundle --ship
```

## Package solution

```bash
gulp package-solution --ship
```

Expected output:

```text
sharepoint/solution/nequette-production-scheduler.sppkg
```