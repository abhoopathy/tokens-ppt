# tokens-ppt

Generates a PowerPoint slide deck from design tokens using [PptxGenJS](https://gitbrent.github.io/PptxGenJS/).

## Setup

**1. Install dependencies**

```bash
npm install
```

**2. Pull the tokens submodule**

```bash
npm run submodule:pull
```

This initializes and pulls `tokens/forge-style-dict` (the `mayorsoffice` branch of [forge-style-dict](https://github.com/abhoopathy/forge-style-dict)).

## Commands

| Command | Description |
|---|---|
| `npm run submodule:pull` | Fetch the latest tokens from the `forge-style-dict` submodule |
| `npm run theme:build` | Package `theme/theme_extracted/` into `theme/mytheme.thmx` |
| `npm run pptx:build` | Generate the slide deck at `slides/output.pptx` |
