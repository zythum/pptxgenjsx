# @zythum02/pptxgenjsx

[English](./README.md) | 中文

一个用于构建 PowerPoint 演示文稿的 JSX 运行时，基于 [pptxgenjs](https://github.com/gitbrent/pptxgenjs)。以 JSX 组件的方式编写幻灯片，并渲染为 `.pptx` 文件。

```tsx
import { Deck, Slide, Text, TextRun, Rect } from "@zythum02/pptxgenjsx";
import { renderPptx } from "@zythum02/pptxgenjsx/render";

await renderPptx(
  <Deck title="My Deck">
    <Slide>
      <Rect x={0} y={0} w={13.333} h={7.5} fill={{ color: "1E1E2E" }} />
      <Text x={1} y={3} w={11} h={1.5}>
        <TextRun options={{ fontSize: 44, color: "FFFFFF", bold: true }}>
          Hello, PowerPoint!
        </TextRun>
      </Text>
    </Slide>
  </Deck>,
  { fileName: "output.pptx" },
);
```

> **Fork 说明**：本包基于 Artifact Kit 的 [pptxgenjs-jsx](https://github.com/artifact-kit/pptxgenjs-jsx) 优秀工作。在原始运行时的基础上扩展了更多功能和调整。

## 目录

- [快速开始](#快速开始)
- [安装](#安装)
- [TypeScript 配置](#typescript-配置)
- [组件参考](#组件参考)
  - [Deck & Presentation](#deck--presentation)
  - [Slide](#slide)
  - [Text & TextRun](#text--textrun)
  - [形状 (Shapes)](#形状-shapes)
  - [图表 (Charts)](#图表-charts)
  - [表格 (Tables)](#表格-tables)
  - [图片与媒体 (Images & Media)](#图片与媒体-images--media)
  - [Group](#group)
  - [Raw（逃生舱）](#raw逃生舱)
  - [Fragment](#fragment)
- [异步组件](#异步组件)
- [Context Hooks](#context-hooks)
  - [useSlideContext](#useslidecontext)
  - [useDeckContext](#usedeckcontext)
  - [useGroupContext](#usegroupcontext)
- [懒加载幻灯片](#懒加载幻灯片)
- [百分比坐标](#百分比坐标)
- [分节与母版](#分节与母版)
- [校验](#校验)
- [渲染](#渲染)
- [许可证](#许可证)

---

## 快速开始

使用官方起始模板 **[pptxgen-ts-starter](https://github.com/zythum/pptxgen-ts-starter)** 快速搭建新项目 — 已预配置 tsconfig、示例幻灯片和构建脚本：

```bash
npx degit zythum/pptxgen-ts-starter my-presentation
cd my-presentation
npm install
npm run build     # 生成 output/presentation.pptx
```

## 安装

```bash
npm install @zythum02/pptxgenjsx
```

对等依赖：

- **[pptxgenjs](https://github.com/gitbrent/pptxgenjs) v4.x** — 已作为依赖自动安装。

## TypeScript 配置

在 `tsconfig.json` 中设置 `jsx` 和 `jsxImportSource`：

```json
{
  "compilerOptions": {
    "jsx": "react-jsx",
    "jsxImportSource": "@zythum02/pptxgenjsx"
  }
}
```

> **注意**：这**不是** React — JSX 转换生成的是 `PptxNode` 对象，而非 DOM 元素。不需要 React 依赖。

---

## 组件参考

### Deck / Presentation

每个演示文稿的根元素。映射到 `new PptxGenJS()`。

```tsx
<Deck>
  <Slide>...</Slide>
</Deck>
```

`<Presentation>` 是 `<Deck>` 的别名。

#### 布局

使用 `layout` 属性控制幻灯片尺寸。支持两种形式：

**1. 内置布局名称**（字符串）— pptxgenjs 提供四种标准预设：

| 名称             | 尺寸          | 宽高比       |
| ---------------- | ------------- | ------------ |
| `"LAYOUT_4x3"`   | 10" × 7.5"    | 4:3          |
| `"LAYOUT_16x9"`  | 10" × 5.625"  | 16:9         |
| `"LAYOUT_16x10"` | 10" × 6.25"   | 16:10        |
| `"LAYOUT_WIDE"`  | 13.33" × 7.5" | 16:9（宽屏） |

默认值：`"LAYOUT_WIDE"`（13.33" × 7.5"）。

```tsx
<Deck layout="LAYOUT_16x9">
  <Slide>...</Slide>
</Deck>
```

**2. 自定义布局**（对象）— 通过 `PresLayout` 对象定义任意尺寸，包含 `name`、`width` 和 `height`（单位为英寸）：

```tsx
<Deck layout={{ name: "A4", width: 10.83, height: 7.82 }}>
  <Slide>...</Slide>
</Deck>
```

对于多个自定义布局，使用 `layouts` 属性（`PresLayout` 数组）：

```tsx
<Deck
  layout="A4"
  layouts={[
    { name: "A4", width: 10.83, height: 7.82 },
    { name: "Letter", width: 10, height: 7.5 },
  ]}
>
  <Slide>...</Slide>
</Deck>
```

第一个匹配的布局名称将成为演示文稿的活动布局。

### Slide

单张幻灯片。映射到 `pptx.addSlide()`。

```tsx
<Slide>
  <Text>A simple slide</Text>
</Slide>
```

**懒加载幻灯片**（参见[懒加载幻灯片](#懒加载幻灯片)）：

```tsx
<Slide component={() => import("./slides/chart-slide")} />
```

### Text & TextRun

**`<Text>`** — 文本框或富文本容器。映射到 `slide.addText()`。

```tsx
// 简单文本（字符串来自 children）
<Text x={1} y={1} w={8} h={1} fontSize={24} color="333333">
  Hello, World!
</Text>

// 包含多个 TextRun 的富文本
<Text x={1} y={2.5} w={8} h={1.5} valign="middle">
  <TextRun options={{ fontSize: 18, color: "666666" }}>普通文本 </TextRun>
  <TextRun options={{ fontSize: 18, color: "0066CC", bold: true }}>
    粗体蓝色
  </TextRun>
</Text>
```

**`<TextRun>`** — `<Text>` 内的单个格式化文本段。`options` 接受 pptxgenjs 的 `TextProps`（fontSize、color、bold、italic、fontFace 等）。文本内容来自字符串 children（或 `text` 属性）：

```tsx
<TextRun options={{ fontSize: 18, color: "0066CC", bold: true }}>粗体蓝色</TextRun>
```

**文本输入模式。** `<Text>` 元素有四种接收内容的方式。当存在多种时，按以下优先级应用（高优先级生效，其余忽略）：

1. 子 `<TextRun />` 节点 — 混合其中的纯字符串/数字 children 被转换为默认样式的 run（保持顺序）
2. `runs` 属性（富文本数组 — 与 pptxgenjs `TextProps[]` 相同结构）
3. `text` 属性
4. 纯字符串 children

纯字符串可与 `<TextRun />` children 自由混合 — 每个字符串片段会原位变成默认样式的 run：

```tsx
<Text x={1} y={4} w={8} h={1}>
  {"普通片段 "}
  <TextRun options={{ color: "0066CC", bold: true }}>高亮</TextRun>
</Text>
```

元素之间的纯空白文本（例如多行 JSX 格式化产生的空白）会被忽略 — 请将显式空格附加到 run 的文本中（`<TextRun>A </TextRun>`）。混合其他模式（例如 `runs` 或 `text` 属性与 children 并存）仍会丢弃低优先级内容。

```tsx
// runs 属性（命令式富文本）
<Text
  x={1}
  y={1}
  w={8}
  h={1}
  runs={[
    { text: "普通 ", options: { fontSize: 18 } },
    { text: "粗体", options: { fontSize: 18, bold: true } },
  ]}
/>
```

**带形状背景的文本。** `shape` 是有效的 pptxgenjs 文本选项，会被转发到 `slide.addText()`：

```tsx
<Text
  x={1}
  y={1}
  w={4}
  h={2}
  shape="roundRect"
  fill={{ color: "EDE9FE" }}
  margin={18} // 文本边距使用磅值（~0.25"），不是英寸
  valign="middle"
>
  带形状的文本框
</Text>
```

> 注意：pptxgenjs 文本的 `margin` 值以**磅**为单位（例如 `18` ≈ 0.25"），不是英寸。

### 形状 (Shapes)

所有 pptxgenjs 形状都可作为 JSX 组件使用。每个组件支持标准定位属性（`x`、`y`、`w`、`h`）以及通过 `options` 或顶层属性传递的形状特定选项。

> **形状是叶元素** — 它们不渲染 children。嵌套内容
>（例如 `<RoundRect><Text>…</Text></RoundRect>`）在 TSX 中是编译时错误，
> 否则是硬运行时错误；永远不会被静默丢弃。
>
> 要在形状上放置文本，使用以下**其一**：
>
> 1. 带形状背景的单个文本框：`<Text shape="roundRect" …>`（当文本和形状共用一个框时推荐 — 参见 [Text & TextRun](#text--textrun)）。
> 2. 在形状上层叠加一个同级 `<Text>`（当文本和形状需要独立样式/阴影/几何形状时）。
> 3. 仅当需要相对坐标系或要一起移动/缩放时才使用 `<Group>`。

```tsx
<Rect x={0} y={0} w={13.333} h={7.5} fill={{ color: "1E1E2E" }} />
<Ellipse x={1} y={1} w={3} h={2} fill={{ color: "FF6B6B" }} />
<Triangle x={5} y={1} w={2} h={2} fill={{ color: "4ECDC4" }} />
<RoundRect x={1} y={4} w={4} h={2} fill={{ color: "45B7D1" }} rectRadius={0.3} />
<Line x={1} y={1} w={5} h={0} line={{ color: "FF0000", width: 2 }} />
<Cloud x={1} y={1} w={3} h={2} fill={{ color: "E8F4FD" }} />
<Heart x={5} y={1} w={2} h={2} fill={{ color: "FF4081" }} />
```

**`<LineBetween>`** — 连接两个绝对坐标的线段。支持百分比坐标。

```tsx
<LineBetween
  x1={0}
  y1={0}
  x2={13.333}
  y2={7.5}
  line={{ color: "999999", width: 1, dashType: "dash" }}
/>
```

可用形状组件：

| 组件                       | pptxgenjs 形状                      |
| -------------------------- | ----------------------------------- |
| `Rect`                     | `rect`                              |
| `RoundRect`                | `roundRect`                         |
| `Ellipse` / `Oval`         | `ellipse`                           |
| `Triangle`                 | `triangle`                          |
| `RightTriangle`            | `rtTriangle`                        |
| `Diamond`                  | `diamond`                           |
| `Pentagon`                 | `pentagon`                          |
| `Hexagon`                  | `hexagon`                           |
| `Star` / `Star5`           | `star5`                             |
| `Star4`                    | `star4`                             |
| `Star6`                    | `star6`                             |
| `Star8`                    | `star8`                             |
| `Star10`                   | `star10`                            |
| `Line`                     | `line`                              |
| `LineBetween`              | `line`（计算边界框）                |
| `Arc`                      | `arc`                               |
| `BlockArc`                 | `blockArc`                          |
| `PieShape`                 | `pie`                               |
| `CustomGeometry`           | `custGeom`                          |
| `LeftArrow` / `RightArrow` | `leftArrow` / `rightArrow`          |
| `UpArrow` / `DownArrow`    | `upArrow` / `downArrow`             |
| `LeftRightArrow`           | `leftRightArrow`                    |
| `UpDownArrow`              | `upDownArrow`                       |
| `Chevron`                  | `chevron`                           |
| `Cloud`                    | `cloud`                             |
| `Heart`                    | `heart`                             |
| `Donut`                    | `donut`                             |
| `Plus`                     | `plus`                              |

### 图表 (Charts)

```tsx
// 通用图表，显式指定 `type`
<Chart
  x={1} y={1} w={10} h={5}
  type="bar"
  data={[
    { name: "Q1", labels: ["Jan", "Feb", "Mar"], values: [100, 150, 200] },
    { name: "Q2", labels: ["Jan", "Feb", "Mar"], values: [120, 180, 160] },
  ]}
  showTitle="Quarterly Sales"
  showLegend={true}
/>

// 或使用类型化图表组件
<BarChart x={1} y={1} w={10} h={5} data={[...]} showValue={true} />
<LineChart x={1} y={1} w={10} h={5} data={[...]} lineSize={3} />
<PieChart x={1} y={1} w={6} h={5} data={[...]} showPercent={true} />
```

可用图表组件：`AreaChart`、`BarChart`、`Bar3DChart`、`BubbleChart`、`DoughnutChart`、`LineChart`、`PieChart`、`RadarChart`、`ScatterChart`。

### 表格 (Tables)

```tsx
<Table x={1} y={1} w={10} h={3} fontSize={12} border={{ type: "solid", color: "CCCCCC" }}>
  <TableRow>
    <TableCell options={{ fill: { color: "4472C4" }, color: "FFFFFF", bold: true }}>名称</TableCell>
    <TableCell options={{ fill: { color: "4472C4" }, color: "FFFFFF", bold: true }}>
      数值
    </TableCell>
  </TableRow>
  <TableRow>
    <TableCell>项目 A</TableCell>
    <TableCell>100</TableCell>
  </TableRow>
  <TableRow>
    <TableCell>项目 B</TableCell>
    <TableCell>200</TableCell>
  </TableRow>
</Table>
```

`TableToSlides` 可将 HTML 表格拆分到多张幻灯片（仅浏览器运行时）。

### 图片与媒体 (Images & Media)

```tsx
<Image
  x={1} y={1} w={5} h={3}
  path="https://example.com/image.png"
  sizing={{ type: "contain", w: 5, h: 3 }}
/>

<Media
  x={1} y={1} w={6} h={4}
  path="https://example.com/video.mp4"
  mediaType="video"
/>
```

### Group

一个逻辑容器，将所有子元素相对于组的位置进行偏移。子坐标相对于组的虚拟画布。

```tsx
<Group x={1} y={1} w={10} h={5}>
  {/* 组内 (0, 0) → 幻灯片上 (1, 1) */}
  <Rect x={0} y={0} w={10} h={5} fill={{ color: "F0F0F0" }} />
  {/* 组内 "50%" → 距组原点 5" → 距幻灯片原点 6" */}
  <Text x="50%" y="50%" w={4} h={1}>
    <TextRun options={{ fontSize: 18 }}>在组内居中</TextRun>
  </Text>
</Group>
```

关键特性：

- **坐标变换**：所有子元素的 `x`、`y`、`w`、`h` 值相对于组的虚拟画布解析。百分比字符串按组的 `w`（x/w 方向）或 `h`（y/h 方向）解析，然后加上组的绝对位置偏移。
- **嵌套组**：组可以嵌套 — 每层累加其偏移量。
- **上下文感知**：子元素可使用 `useGroupContext()` 获取组的虚拟画布尺寸。

### Raw（逃生舱）

用于 pptxgenjs 中未被专用组件覆盖的功能。

```tsx
<Raw
  render={({ pptx, slide, node }) => {
    // 直接访问 slide.addShape()、slide.addText() 等
    slide.addShape("rect", { x: 1, y: 1, w: 5, h: 3, fill: { color: "FF0000" } });
  }}
/>
```

> Raw 的 children **不会**自动渲染 — 它们只能通过 `render` 回调内的
> `context.node.children` 访问。使用 props 进行配置；仅当回调自行读取时才使用 children。

### Fragment

将多个子元素分组而不产生包装元素。当组件需要返回多个同级元素时很有用（例如在懒加载幻灯片中或 map 内部）。

**简写形式 `<>...</>`** — 用于不需要 props 的简单分组：

```tsx
// slides/title-slide.tsx
export default function TitleSlide() {
  return (
    <>
      <Text x={1} y={3} w={10} h={1.5} fontSize={44} bold>
        Welcome
      </Text>
      <Text x={1} y={4.5} w={10} h={1} fontSize={18} color="666666">
        Subtitle text
      </Text>
    </>
  );
}
```

**显式 `<Fragment>`** — 当需要 `key` 属性时（例如在 `.map()` 循环中）：

```tsx
import { Fragment } from "@zythum02/pptxgenjsx";

<Slide>
  {items.map((item) => (
    <Fragment key={item.id}>
      <Text x={1} y={item.y}>
        {item.name}
      </Text>
      <Text x={5} y={item.y}>
        {item.value}
      </Text>
    </Fragment>
  ))}
</Slide>;
```

> **注意**：`<>...</>` 是 JSX 编译时语法 — 不支持 `key` 等 props。对于动态列表，请始终使用 `<Fragment key={...}>`。

---

## 异步组件

组件可以是 `async` 函数 — 渲染时会自动检测并懒解析：

```tsx
// slides/data-slide.tsx
export default async function DataSlide() {
  const res = await fetch("https://api.example.com/data");
  const data = await res.json();

  return (
    <Slide>
      <Text x={1} y={1} w={8} h={1} fontSize={32} bold>
        {data.title}
      </Text>
      <Text x={1} y={2.5} w={8} h={4} fontSize={16}>
        {data.description}
      </Text>
    </Slide>
  );
}
```

```tsx
// main.tsx
await renderPptx(
  <Deck>
    <Slide>{/* ... */}</Slide>
    <DataSlide /> {/* async — 自动解析 */}
  </Deck>,
  { fileName: "output.pptx" },
);
```

这之所以可行，是因为 JSX 工厂（`jsx`）将异步组件结果包装在 `PptxNodePromise` 中，渲染器在树遍历时解析它们。

---

## Context Hooks

Context hooks 提供当前渲染环境的运行时信息。

**工作原理**：当 TypeScript 编译 JSX 时，组件工厂在 JSX 构建期间**不会**被调用。它们被包装在延迟节点中，在渲染期间才执行 — 此时渲染器已通过 `AsyncLocalStorage` 建立了上下文存储。这就是为什么你可以编写调用 `useSlideContext()` 的组件作为 `<Slide>` 的直接子元素，即使 JSX 看起来是"急切"构建的。

### useSlideContext

暴露当前幻灯片的索引和总数。

```tsx
import { useSlideContext } from "@zythum02/pptxgenjsx";

function SlideNumber() {
  const { index, total, sectionTitle } = useSlideContext();

  return (
    <Text x={1} y={6.5} w={10} h={0.5} fontSize={10} color="999999">
      第 {index} 页，共 {total} 页
      {sectionTitle ? ` · ${sectionTitle}` : ""}
    </Text>
  );
}
```

```tsx
// 在 Slide 中使用：
<Slide>
  {/* ... 幻灯片内容 ... */}
  <SlideNumber />
</Slide>
```

### useDeckContext

暴露演示文稿的幻灯片尺寸（宽度、高度，单位英寸）。

```tsx
import { useDeckContext } from "@zythum02/pptxgenjsx";

function FullBleedBackground() {
  const { width, height } = useDeckContext();
  return <Rect x={0} y={0} w={width} h={height} fill={{ color: "1E1E2E" }} />;
}
```

### useGroupContext

暴露当前组的绝对偏移和虚拟画布尺寸。在 `<Group>` 外调用时，回退到演示文稿尺寸（偏移为零）。

```tsx
import { useGroupContext } from "@zythum02/pptxgenjsx";

function ProgressBar() {
  const { width } = useGroupContext();
  return <Rect x={0} y={0} w={width * 0.7} h={0.4} fill={{ color: "4CAF50" }} />;
}
```

---

## 懒加载幻灯片

使用 `<Slide>` 上的 `component` 属性延迟加载幻灯片定义 — 类似于 React Router 的懒路由加载。

```tsx
// slides/title-slide.tsx
export default function TitleSlide() {
  return (
    <>
      <Text x={1} y={3} w={10} h={1.5} fontSize={44} bold>
        Welcome
      </Text>
    </>
  );
}
```

```tsx
// main.tsx
<Slide component={() => import("./slides/title-slide")} />
```

组件的返回值渲染在声明 `component` 的 `<Slide>` **内部**，因此它应该只提供幻灯片**内容** — 而不是另一个 `<Slide>` 元素。Context hooks 在懒加载组件内部正常工作。

---

## 百分比坐标

`x`、`y`、`w`、`h` 值可以指定为百分比字符串（例如 `"50%"`、`"100%"`），相对于封闭上下文解析：

- **在 `<Group>` 内部**：按组的 `w`（x/w 方向）或 `h`（y/h 方向）解析。
- **直接在 `<Slide>` 内部**：按演示文稿布局的幻灯片尺寸解析。

```tsx
<Group x={1} y={1} w={10} h={5}>
  {/* 组宽度的 50%（5"），组高度的 25%（1.25"） */}
  <Rect x="25%" y="25%" w="50%" h="50%" fill={{ color: "4ECDC4" }} />
</Group>
```

这适用于所有定位属性：所有形状/文本/图片组件的 `x`、`y`、`w`、`h`，以及 `LineBetween` 的 `x1`、`y1`、`x2`、`y2`。

---

## 分节与母版

### 分节 (Sections)

将幻灯片分组到 PowerPoint 大纲视图中的命名分节。

```tsx
<Deck>
  <Slide>...</Slide>
  <Section title="概览">
    <Slide>...</Slide>
    <Slide>...</Slide>
  </Section>
  <Section title="详情">
    <Slide>...</Slide>
  </Section>
</Deck>
```

### 母版 (Masters)

使用可复用的布局对象定义幻灯片母版。

```tsx
<Deck>
  <Master name="myMaster" background={{ fill: "F5F5F5" }}>
    <Text x={1} y={0.3} w={10} h={0.5} fontSize={10} color="999999">
      机密
    </Text>
    <Rect x={0} y={7} w={13.333} h={0.5} fill={{ color: "4472C4" }} />
  </Master>
  <Slide masterName="myMaster">...</Slide>
</Deck>
```

你也可以在母版中使用 `<Placeholder>`：

```tsx
<Master name="content">
  <Placeholder options={{ name: "Body", type: "body", x: 1, y: 1, w: 10, h: 5 }} />
</Master>
```

## 校验

`validateDeck()` 函数在渲染前检查幻灯片树中的常见错误。它是**异步的** — 始终使用 `await`：

```tsx
import { validateDeck } from "@zythum02/pptxgenjsx/render";

const deck = <Deck>{/* ... */}</Deck>;
const issues = await validateDeck(deck);

if (issues.length > 0) {
  for (const issue of issues) {
    console.error(`[${issue.level}] ${issue.message}`);
  }
}
```

校验可捕获：

- 无效的子类型（例如 `<Text>` 直接放在 `<Deck>` 内）
- **带 children 的叶组件**（`child.leaf`，例如 `<RoundRect>` 包含 `<Text>`）
- 缺少必需的 props（例如 `CustomGeometry` 没有 `points`）
- 可疑的 prop 用法（例如 `RoundRect` 上的 `angleRange`）
- 混合的 `<Text>` 输入模式（`text.input.mixed`）— 当低优先级来源会被忽略时发出警告，并推荐使用 `<TextRun />` children 实现富文本
- 无效的 `LineBetween` 端点

---

## 渲染

### 输出到文件

```tsx
import { renderPptx } from "@zythum02/pptxgenjsx/render";

await renderPptx(<Deck>{/* ... */}</Deck>, {
  fileName: "output/presentation.pptx",
});
```

还支持额外的 pptxgenjs `writeFile` 选项：

```tsx
await renderPptx(<Deck>{/* ... */}</Deck>, {
  fileName: "output.pptx",
  compression: true, // 启用 ZIP 压缩
  zipOptions: { level: 9 }, // 压缩级别
});
```

### 输出到 Buffer / Blob

```tsx
import { writePptx } from "@zythum02/pptxgenjsx/render";

const buffer = await writePptx(<Deck>{/* ... */}</Deck>, {
  outputType: "arraybuffer", // "arraybuffer" | "blob" | "uint8array" | "base64" | "nodebuffer"
});
```

### 复用现有 PptxGenJS 实例

```tsx
import PptxGenJS from "pptxgenjs";

const pptx = new PptxGenJS();
// ... 配置实例 ...
await renderPptx(<Deck>{/* ... */}</Deck>, { pptx });
```

### 别名

```tsx
import { render, write } from "@zythum02/pptxgenjsx/render";

await render(<Deck>{/* ... */}</Deck>, { fileName: "output.pptx" });
```

## 许可证

MIT
