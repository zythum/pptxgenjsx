# @zythum02/pptxgenjsx

A JSX runtime for building PowerPoint presentations with [pptxgenjs](https://github.com/gitbrent/pptxgenjs). Write your slides as JSX components and render them to `.pptx` files.

> **Fork notice**: This package is based on the excellent work of [pptxgenjs-jsx](https://github.com/artifact-kit/pptxgenjs-jsx) by Artifact Kit. It extends the original runtime with additional features and adjustments.

> 💡 **Recommended**: Quickly scaffold a new project with the official starter template
> **[pptxgen-ts-starter](https://github.com/zythum/pptxgen-ts-starter)** — preconfigured with tsconfig, sample deck, and build script.

## Features

- **JSX syntax** — Declare slides, text, shapes, and charts with familiar JSX
- **Async components** — Components can be `async`, allowing data fetching directly inside slide definitions
- **Full pptxgenjs coverage** — All shapes and chart types are supported
- **TypeScript first** — Full type definitions included

## Usage

### 1. Configure `tsconfig.json`

```json
{
  "compilerOptions": {
    "jsx": "react-jsx",
    "jsxImportSource": "@zythum02/pptxgenjsx"
  }
}
```

### 2. Write your deck (e.g. `src/ppt.tsx`)

```tsx
import { Deck, Slide, Text, TextRun, Rect } from "@zythum02/pptxgenjsx";
import { renderPptx } from "@zythum02/pptxgenjsx/render";

// Synchronous slide — static content
function TitleSlide() {
  return (
    <Slide>
      <Rect x={0} y={0} w={13.333} h={7.5} fill={{ color: "1E1E2E" }} />
      <Text x={1} y={3} w={11} h={1.5}>
        <TextRun options={{ fontSize: 44, color: "FFFFFF", bold: true }}>My Presentation</TextRun>
      </Text>
    </Slide>
  );
}

// Async slide — fetch data directly inside the component
async function DataSlide() {
  const res = await fetch("https://api.example.com/data");
  const data = await res.json();

  return (
    <Slide>
      <Rect x={0} y={0} w={13.333} h={7.5} fill={{ color: "FFFFFF" }} />
      <Text x={1} y={1} w={10} h={1}>
        <TextRun options={{ fontSize: 32, bold: true }}>{data.title}</TextRun>
      </Text>
    </Slide>
  );
}

async function main() {
  await renderPptx(
    <Deck>
      <TitleSlide />
      <DataSlide />
    </Deck>,
    { fileName: "output/presentation.pptx" },
  );
}

main();
```

## License

MIT
