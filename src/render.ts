/**
 * Rendering engine for pptxgenjsx.
 *
 * All tree-traversal functions are async to support PptxNodePromise resolution.
 * Component factories and type definitions are in `./components.ts`.
 */

import PptxGenJS from "pptxgenjs";
import type PptxGenJSType from "pptxgenjs";

import {
  type PptxNode,
  type PptxElement,
  type PptxPresentation,
  type PptxSlide,
  isPptxNode,
  resolveChild,
} from "./node.js";

import {
  type SlideContextInfo,
  type DeckContextInfo,
  type GroupContextInfo,
  useGroupContext,
  withDeckContext,
  withSlideContext,
  withGroupContext,
} from "./context.js";

// ════════════════════════════════════════════════════════════════════
// RENDERING ENGINE
// ════════════════════════════════════════════════════════════════════

// ── Reserved prop keys (not passed to pptxgenjs) ─────────────────

const RESERVED_PROPS = new Set([
  "children",
  "options",
  "text",
  "runs",
  "shape",
  "data",
  "rows",
  "cells",
  "eleId",
  "render",
  "component",
]);

// ── Shape type sets ───────────────────────────────────────────────

const CHART_TYPES = new Set([
  "Chart",
  "AreaChart",
  "BarChart",
  "Bar3DChart",
  "BubbleChart",
  "DoughnutChart",
  "LineChart",
  "PieChart",
  "RadarChart",
  "ScatterChart",
]);

const SHAPE_TYPES = new Set([
  "Shape",
  "Rect",
  "RoundRect",
  "Ellipse",
  "Oval",
  "Triangle",
  "RightTriangle",
  "Diamond",
  "Pentagon",
  "Hexagon",
  "Star",
  "Star4",
  "Star5",
  "Star6",
  "Star8",
  "Star10",
  "Arc",
  "BlockArc",
  "PieShape",
  "CustomGeometry",
  "LeftArrow",
  "RightArrow",
  "UpArrow",
  "DownArrow",
  "LeftRightArrow",
  "UpDownArrow",
  "Chevron",
  "Cloud",
  "Heart",
  "Donut",
  "Plus",
]);

// ── Public API ────────────────────────────────────────────────────

export interface CreatePptxOptions {
  /** Reuse an existing PptxGenJS instance instead of creating a new one. */
  pptx?: PptxGenJSType;
}

export type RenderPptxOptions = PptxGenJSType.WriteFileProps & CreatePptxOptions;
export type WritePptxOptions = PptxGenJSType.WriteProps & CreatePptxOptions;

/**
 * Create a PptxGenJS instance from a Deck or Slide node tree.
 * Resolves any PptxNodePromise wrappers during traversal.
 */
export async function createPptx(
  root: PptxElement,
  options: CreatePptxOptions = {},
): Promise<PptxGenJSType> {
  const node = await resolveChild(root);
  const pptx = options.pptx ?? new PptxGenJS();

  if (node.type === "Deck") {
    applyDeckProps(pptx, node.props as any);

    const total = countDeckSlides(node);
    const width = (pptx as any).width ?? 13.333;
    const height = (pptx as any).height ?? 7.5;

    const ctx: SlideRenderContext = {
      pptx,
      slideIndex: 0,
      total,
      width,
      height,
    };

    const deckCtx: DeckContextInfo = { width, height };

    await withDeckContext(deckCtx, () => processDeckChildren(node, ctx));
    return pptx;
  }

  if (node.type === "Slide") {
    const width = (pptx as any).width ?? 13.333;
    const height = (pptx as any).height ?? 7.5;

    const ctx: SlideRenderContext = {
      pptx,
      slideIndex: 0,
      total: 1,
      width,
      height,
    };

    const deckCtx: DeckContextInfo = { width, height };

    await withDeckContext(deckCtx, () => processSlide(node, ctx));
    return pptx;
  }

  throw new Error(`createPptx expected a Deck or Slide root, got ${node.type}.`);
}

// ── Slide counting ────────────────────────────────────────────────

/** Count all `<Slide>` nodes in the node tree recursively. */
function countDeckSlides(node: PptxNode): number {
  if (node.type === "Slide") return 1;

  let count = 0;
  for (const rawChild of node.children) {
    if (!isPptxNode(rawChild)) continue;
    count += countDeckSlides(rawChild);
  }
  return count;
}

/**
 * Render to a .pptx file.
 * Returns the file name (as per PptxGenJS.writeFile).
 */
export async function renderPptx(
  root: PptxElement,
  options: RenderPptxOptions = {},
): Promise<string> {
  const { pptx, ...writeOptions } = options;
  const instance = await createPptx(root, { pptx });
  return instance.writeFile(writeOptions);
}

/**
 * Write to a buffer / ArrayBuffer / Blob.
 */
export async function writePptx(
  root: PptxElement,
  options: WritePptxOptions = {},
): Promise<string | ArrayBuffer | Blob | Uint8Array> {
  const { pptx, ...writeOptions } = options;
  const instance = await createPptx(root, { pptx });
  return instance.write(writeOptions);
}

/** Alias for renderPptx. */
export const render = renderPptx;
/** Alias for writePptx. */
export const write = writePptx;

// ── Deck props ────────────────────────────────────────────────────

function applyDeckProps(pptx: PptxPresentation, props: Record<string, any>): void {
  const {
    layouts,
    sections,
    masters,
    layout,
    title,
    author,
    company,
    subject,
    revision,
    rtlMode,
    theme,
  } = props;

  layouts?.forEach((l: any) => pptx.defineLayout(l));
  sections?.forEach((s: any) => pptx.addSection(s));
  masters?.forEach((m: any) => pptx.defineSlideMaster(m));

  if (typeof layout === "string") {
    pptx.layout = layout;
  } else if (layout) {
    pptx.defineLayout(layout);
    pptx.layout = layout.name;
  }

  setIfDefined(pptx, "title", title);
  setIfDefined(pptx, "author", author);
  setIfDefined(pptx, "company", company);
  setIfDefined(pptx, "subject", subject);
  setIfDefined(pptx, "revision", revision);
  setIfDefined(pptx, "rtlMode", rtlMode);
  setIfDefined(pptx, "theme", theme);
}

// ── Deck children traversal (async) ───────────────────────────────

/** Internal rendering context that carries slide counting state. */
interface SlideRenderContext {
  pptx: PptxPresentation;
  slideIndex: number;
  total: number;
  width: number;
  height: number;
}

async function processDeckChildren(node: PptxNode, ctx: SlideRenderContext): Promise<void> {
  for (const rawChild of node.children) {
    const child = await resolveChild(rawChild);
    if (!isPptxNode(child)) continue;

    switch (child.type) {
      case "Layout":
        ctx.pptx.defineLayout(child.props as any);
        break;
      case "Master":
        ctx.pptx.defineSlideMaster(buildMaster(child) as any);
        break;
      case "Section":
        await processSection(child, ctx);
        break;
      case "Slide":
        await processSlide(child, ctx);
        break;
      case "TableToSlides":
        processTableToSlides(child, ctx);
        break;
      case "Raw":
        (child.props as any).render({ pptx: ctx.pptx, node: child });
        break;
      case "Fragment":
        await processDeckChildren(child, ctx);
        break;
      default:
        throw new Error(`${child.type} must be inside a Slide, Section, or Master.`);
    }
  }
}

// ── Section ───────────────────────────────────────────────────────

async function processSection(node: PptxNode, ctx: SlideRenderContext): Promise<void> {
  const props = node.props as any;
  ctx.pptx.addSection({ title: props.title, order: props.order });

  for (const rawChild of node.children) {
    const child = await resolveChild(rawChild);
    if (!isPptxNode(child)) continue;

    if (child.type !== "Slide") {
      throw new Error(`Section children must be Slide nodes. Got ${child.type}.`);
    }
    await processSlide(child, ctx, props.title);
  }
}

// ── Slide ─────────────────────────────────────────────────────────

async function processSlide(
  node: PptxNode,
  ctx: SlideRenderContext,
  sectionTitle?: string,
): Promise<void> {
  const props = node.props as any;

  // Increment and capture the 1-based slide index
  ctx.slideIndex += 1;
  const index = ctx.slideIndex;
  const total = ctx.total;

  const resolvedSectionTitle = props.sectionTitle ?? sectionTitle;

  const slide = ctx.pptx.addSlide({
    masterName: props.masterName,
    sectionTitle: resolvedSectionTitle,
  });

  setIfDefined(slide, "background", props.background);
  setIfDefined(slide, "color", props.color);
  setIfDefined(slide, "hidden", props.hidden);
  setIfDefined(slide, "slideNumber", props.slideNumber);

  const slideCtx: SlideContextInfo = {
    index,
    total,
    sectionTitle: resolvedSectionTitle,
  };

  // ── Lazy `component` prop — like React Router's `component={() => import('./path')}` ──
  if (props.component) {
    const module = await props.component();
    const LazyComponent = module.default;

    // Wrap only the factory call in context — sync variable approach is safe
    // because the factory's JSX-building phase executes synchronously.
    const lazyNode = await withSlideContext(slideCtx, () => resolveChild(LazyComponent({})));
    await renderSlideElement(lazyNode, { pptx: ctx.pptx, slide });
    return;
  }

  // ── Direct children ──
  await withSlideContext(slideCtx, async () => {
    for (const rawChild of node.children) {
      const child = await resolveChild(rawChild);
      if (!isPptxNode(child)) continue;
      await renderSlideElement(child, { pptx: ctx.pptx, slide });
    }
  });
}

// ── Slide element rendering ───────────────────────────────────────

async function renderSlideElement(
  node: PptxNode,
  ctx: { pptx: PptxPresentation; slide: PptxSlide },
): Promise<void> {
  const { slide } = ctx;

  if (CHART_TYPES.has(node.type)) {
    renderChart(node, slide);
    return;
  }

  if (SHAPE_TYPES.has(node.type)) {
    renderShape(node, slide);
    return;
  }

  switch (node.type) {
    case "Text":
      renderText(node, slide);
      break;
    case "Line":
      renderLine(node, slide);
      break;
    case "LineBetween":
      renderLineBetween(node, slide);
      break;
    case "Image":
      renderImage(node, slide);
      break;
    case "Media":
      renderMedia(node, slide);
      break;
    case "Table":
      renderTable(node, slide);
      break;
    case "Notes":
      slide.addNotes(getTextContent(node, (node.props as any).text));
      break;
    case "Raw":
      (node.props as any).render({ pptx: ctx.pptx, slide, node });
      break;
    case "Group": {
      const props = node.props as any;
      const parentGroup = useGroupContext(); // deck fallback when not inside another group
      const absX = resolveCoord(props.x, parentGroup.width, 0) + parentGroup.x;
      const absY = resolveCoord(props.y, parentGroup.height, 0) + parentGroup.y;
      const grpW = resolveCoord(props.w, parentGroup.width, parentGroup.width);
      const grpH = resolveCoord(props.h, parentGroup.height, parentGroup.height);

      const groupInfo: GroupContextInfo = { x: absX, y: absY, width: grpW, height: grpH };

      await withGroupContext(groupInfo, async () => {
        for (const rawChild of node.children) {
          const child = await resolveChild(rawChild);
          if (isPptxNode(child)) {
            await renderSlideElement(child, ctx);
          }
        }
      });
      break;
    }
    case "Fragment":
      for (const rawChild of node.children) {
        const child = await resolveChild(rawChild);
        if (isPptxNode(child)) {
          await renderSlideElement(child, ctx);
        }
      }
      break;
    default:
      throw new Error(`${node.type} cannot be rendered directly on a slide.`);
  }
}

// ── Coordinate helpers ─────────────────────────────────────────────

/**
 * Resolve a pptxgenjs Coord (number or `${number}%`) to an absolute
 * number using a reference dimension (e.g. group width for x / w,
 * group height for y / h).
 *
 * - `5`       → 5
 * - `"50%"`   → `0.5 * reference`
 * - `undefined` / `null` → `fallback` (default 0)
 */
function resolveCoord(
  value: number | `${number}%` | undefined | null,
  reference: number,
  fallback: number = 0,
): number {
  if (value == null) return fallback;
  if (typeof value === "number") return value;
  // `${number}%` — e.g. "50%"
  const pct = parseFloat(value);
  if (isNaN(pct)) return fallback;
  if (!isFinite(pct)) return fallback;
  return (pct / 100) * reference;
}

/**
 * Apply the current group's absolute x/y offset to an element's props.
 * Percentage-based Coord values (e.g. `"50%"`) are resolved relative to
 * the group's virtual canvas **before** adding the group offset.
 * When no group is active (i.e., direct slide children), returns the
 * props unchanged — percentages pass through to pptxgenjs natively.
 */
function offsetXY<T extends Record<string, any>>(props: T): T {
  const { x: groupX, y: groupY, width: grpW, height: grpH } = useGroupContext();
  if (
    groupX === 0 &&
    groupY === 0 &&
    typeof props.x !== "string" &&
    typeof props.y !== "string" &&
    typeof props.w !== "string" &&
    typeof props.h !== "string"
  ) {
    return { ...props };
  }
  const resolvedX = resolveCoord(props.x, grpW, 0);
  const resolvedY = resolveCoord(props.y, grpH, 0);
  const resolvedW = resolveCoord(props.w, grpW, props.w);
  const resolvedH = resolveCoord(props.h, grpH, props.h);
  return { ...props, x: resolvedX + groupX, y: resolvedY + groupY, w: resolvedW, h: resolvedH };
}

// ── Individual element renderers ──────────────────────────────────

function renderText(node: PptxNode, slide: PptxSlide): void {
  const props = offsetXY(node.props as any);
  const textRuns = collectTextRuns(node) ?? props.text ?? getTextContent(node);
  slide.addText(textRuns, mergeOptions(props.options, props));
}

function renderShape(node: PptxNode, slide: PptxSlide): void {
  const props = offsetXY(node.props as any);
  slide.addShape(props.shape, mergeOptions(props.options, props, ["shape"]));
}

function renderLine(node: PptxNode, slide: PptxSlide): void {
  const props = offsetXY(node.props as any);
  slide.addShape("line", mergeOptions(props.options, props));
}

function renderLineBetween(node: PptxNode, slide: PptxSlide): void {
  const { x: ox, y: oy, width: grpW, height: grpH } = useGroupContext();
  const props = node.props as any;
  const x1 = resolveCoord(props.x1, grpW, 0) + ox;
  const y1 = resolveCoord(props.y1, grpH, 0) + oy;
  const x2 = resolveCoord(props.x2, grpW, 0) + ox;
  const y2 = resolveCoord(props.y2, grpH, 0) + oy;
  const x = Math.min(x1, x2);
  const y = Math.min(y1, y2);
  const w = Math.abs(x2 - x1);
  const h = Math.abs(y2 - y1);

  slide.addShape("line", {
    ...mergeOptions(props.options, props, ["x1", "y1", "x2", "y2"]),
    x,
    y,
    w,
    h,
    flipH: x2 < x1,
    flipV: y2 < y1,
  });
}

function renderImage(node: PptxNode, slide: PptxSlide): void {
  const props = offsetXY(node.props as any);
  slide.addImage(mergeOptions(props.options, props));
}

function renderMedia(node: PptxNode, slide: PptxSlide): void {
  const props = offsetXY(node.props as any);
  slide.addMedia(mergeOptions(props.options, props) as any);
}

function renderChart(node: PptxNode, slide: PptxSlide): void {
  const props = offsetXY(node.props as any);
  slide.addChart(
    props.type,
    deepClone(props.data),
    mergeOptions(props.options, props, ["type", "data"]),
  );
}

function renderTable(node: PptxNode, slide: PptxSlide): void {
  const props = offsetXY(node.props as any);
  const rows =
    props.rows ??
    node.children
      .filter(isPptxNode)
      .filter((c) => c.type === "TableRow")
      .map(extractTableRow);
  slide.addTable(deepClone(rows), mergeOptions(props.options, props));
}

function processTableToSlides(node: PptxNode, ctx: { pptx: PptxPresentation }): void {
  const props = node.props as any;
  ctx.pptx.tableToSlides(props.eleId, mergeOptions(props.options, props, ["eleId"]));
}

// ── Table helpers ─────────────────────────────────────────────────

function extractTableRow(node: PptxNode): any[] {
  if (node.type !== "TableRow") {
    throw new Error(`Table children must be TableRow nodes. Got ${node.type}.`);
  }
  const props = node.props as any;
  return (
    props.cells ??
    node.children
      .filter(isPptxNode)
      .filter((c) => c.type === "TableCell")
      .map(extractTableCell)
  );
}

function extractTableCell(node: PptxNode): any {
  if (node.type !== "TableCell") {
    throw new Error(`TableRow children must be TableCell nodes. Got ${node.type}.`);
  }
  const props = node.props as any;
  return {
    text: deepClone(props.text ?? getTextContent(node)),
    options: mergeOptions(props.options, props),
  };
}

// ── Master helpers ────────────────────────────────────────────────

function buildMaster(node: PptxNode): Record<string, any> {
  const props = node.props as any;
  const { children, objects = [], ...rest } = props;
  const childObjects = children.filter(isPptxNode).map(extractMasterObject);
  return {
    ...rest,
    objects: deepClone([...objects, ...childObjects]),
  };
}

function extractMasterObject(node: PptxNode): Record<string, any> {
  if (SHAPE_TYPES.has(node.type)) {
    const props = node.props as any;
    const opts = mergeOptions(props.options, props, ["shape"]);
    if (props.shape === "rect") return { rect: opts };
    if (props.shape === "line") return { line: opts };
    throw new Error(
      "Master only supports rect and line shape objects directly. Use the objects prop or Raw for other master shapes.",
    );
  }

  switch (node.type) {
    case "Text": {
      const props = node.props as any;
      return {
        text: {
          text: getTextContent(node, typeof props.text === "string" ? props.text : undefined),
          options: mergeOptions(props.options, props),
        },
      };
    }
    case "Image":
      return { image: mergeOptions((node.props as any).options, node.props) };
    case "Line":
      return { line: mergeOptions((node.props as any).options, node.props) };
    case "LineBetween": {
      const props = node.props as any;
      const x = Math.min(props.x1, props.x2);
      const y = Math.min(props.y1, props.y2);
      const w = Math.abs(props.x2 - props.x1);
      const h = Math.abs(props.y2 - props.y1);
      return {
        line: {
          ...mergeOptions(props.options, props, ["x1", "y1", "x2", "y2"]),
          x,
          y,
          w,
          h,
          flipH: props.x2 < props.x1,
          flipV: props.y2 < props.y1,
        },
      };
    }
    case "Placeholder": {
      const props = node.props as any;
      return {
        placeholder: {
          options: props.options,
          text: getTextContent(node, props.text),
        },
      };
    }
    default:
      throw new Error(
        `${node.type} is not supported inside Master. Use the objects prop or Raw for advanced master content.`,
      );
  }
}

// ── Text helpers ──────────────────────────────────────────────────

function collectTextRuns(node: PptxNode): any[] | undefined {
  const runs = node.children
    .filter(isPptxNode)
    .filter((c) => c.type === "TextRun")
    .map((c) => deepClone(c.props));
  return runs.length > 0 ? runs : undefined;
}

function getTextContent(node: PptxNode, fallback?: string): string {
  if (fallback !== undefined) return fallback;
  return node.children
    .filter((c): c is string | number => typeof c === "string" || typeof c === "number")
    .map(String)
    .join("");
}

// ── Props / options utilities ─────────────────────────────────────

/**
 * Merge component top-level props with an `options` object.
 * Top-level props that match pptxgenjs option keys override `options`.
 * Reserved/structural props (children, options, etc.) are excluded.
 */
function mergeOptions(
  options: Record<string, any> | undefined,
  props: Record<string, any>,
  exclude: string[] = [],
): Record<string, any> {
  const result: Record<string, any> = deepClone(options ?? {});
  const skip = new Set(exclude);

  for (const [key, value] of Object.entries(props)) {
    if (!RESERVED_PROPS.has(key) && !skip.has(key) && value !== undefined) {
      result[key] = deepClone(value);
    }
  }

  return result;
}

/**
 * Deep-clone plain objects and arrays.
 * Preserves PptxGenJS class instances and other non-plain objects.
 */
function deepClone<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map((item) => deepClone(item)) as unknown as T;
  }
  if (value && typeof value === "object") {
    const proto = Object.getPrototypeOf(value);
    if (proto === Object.prototype || proto === null) {
      const result: Record<string, any> = {};
      for (const [key, val] of Object.entries(value as Record<string, any>)) {
        result[key] = deepClone(val);
      }
      return result as T;
    }
  }
  return value;
}

function setIfDefined(obj: any, key: string, value: any): void {
  if (value !== undefined) {
    obj[key] = value;
  }
}

// ════════════════════════════════════════════════════════════════════
// VALIDATION
// ════════════════════════════════════════════════════════════════════

export interface ValidationIssue {
  code: string;
  level: "error" | "warning";
  path: string;
  message: string;
}

// ── Valid child sets ──────────────────────────────────────────────

const DECK_CHILDREN = new Set([
  "Layout",
  "Master",
  "Section",
  "Slide",
  "TableToSlides",
  "Raw",
  "Fragment",
]);

const SECTION_CHILDREN = new Set(["Slide", "Fragment"]);

const SLIDE_CHILDREN = new Set([
  "Group",
  "Text",
  "Shape",
  "Rect",
  "RoundRect",
  "Ellipse",
  "Oval",
  "Triangle",
  "RightTriangle",
  "Diamond",
  "Pentagon",
  "Hexagon",
  "Star",
  "Star4",
  "Star5",
  "Star6",
  "Star8",
  "Star10",
  "Line",
  "LineBetween",
  "Arc",
  "BlockArc",
  "PieShape",
  "CustomGeometry",
  "LeftArrow",
  "RightArrow",
  "UpArrow",
  "DownArrow",
  "LeftRightArrow",
  "UpDownArrow",
  "Chevron",
  "Cloud",
  "Heart",
  "Donut",
  "Plus",
  "Image",
  "Media",
  "Chart",
  "AreaChart",
  "BarChart",
  "Bar3DChart",
  "BubbleChart",
  "DoughnutChart",
  "LineChart",
  "PieChart",
  "RadarChart",
  "ScatterChart",
  "Table",
  "Notes",
  "Raw",
  "Fragment",
]);

const TABLE_CHILDREN = new Set(["TableRow", "Fragment"]);
const TABLE_ROW_CHILDREN = new Set(["TableCell", "Fragment"]);

function getValidChildTypes(type: string): Set<string> | undefined {
  switch (type) {
    case "Deck":
      return DECK_CHILDREN;
    case "Section":
      return SECTION_CHILDREN;
    case "Slide":
      return SLIDE_CHILDREN;
    case "Table":
      return TABLE_CHILDREN;
    case "TableRow":
      return TABLE_ROW_CHILDREN;
    default:
      return undefined;
  }
}

// ── Validation entry point ────────────────────────────────────────

/**
 * Validate a Deck or Slide tree for common mistakes.
 * Returns a list of issues (errors and warnings).
 */
export function validateDeck(root: PptxNode): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  if (root.type !== "Deck" && root.type !== "Slide") {
    issues.push({
      code: "root.invalid",
      level: "error",
      path: root.type,
      message: `Root must be Deck or Slide, got ${root.type}.`,
    });
  }

  walk(root, root.type);
  return issues;

  function walk(node: PptxNode, path: string): void {
    validateNodeProps(node, path, issues);
    for (const child of node.children) {
      if (!isPptxNode(child)) continue;
      validateChildType(node, child, `${path}/${child.type}`, issues);
      walk(child, `${path}/${child.type}`);
    }
  }
}

function validateChildType(
  parent: PptxNode,
  child: PptxNode,
  path: string,
  issues: ValidationIssue[],
): void {
  const validChildren = getValidChildTypes(parent.type);
  if (validChildren && !validChildren.has(child.type)) {
    issues.push({
      code: "child.invalid",
      level: "error",
      path,
      message: `${child.type} is not a valid child of ${parent.type}.`,
    });
  }
}

function validateNodeProps(node: PptxNode, path: string, issues: ValidationIssue[]): void {
  const props = node.props as any;

  if (node.type === "RoundRect" && "angleRange" in props) {
    issues.push({
      code: "shape.prop.invalid",
      level: "warning",
      path,
      message: "RoundRect ignores angleRange; use Arc, PieShape, or BlockArc for angle ranges.",
    });
  }

  if ((node.type === "Arc" || node.type === "PieShape") && "arcThicknessRatio" in props) {
    issues.push({
      code: "shape.prop.invalid",
      level: "warning",
      path,
      message: `${node.type} ignores arcThicknessRatio; use BlockArc for arc thickness.`,
    });
  }

  if (node.type === "LineBetween") {
    for (const key of ["x1", "y1", "x2", "y2"] as const) {
      const val = props[key];
      if (
        val == null ||
        (typeof val !== "number" && typeof val !== "string") ||
        (typeof val === "number" && Number.isNaN(val))
      ) {
        issues.push({
          code: "line.endpoint.invalid",
          level: "error",
          path,
          message: `LineBetween requires ${key} to be a Coord (number or percentage string).`,
        });
      }
    }
  }

  if (
    node.type === "CustomGeometry" &&
    (!Array.isArray(props.points) || props.points.length === 0)
  ) {
    issues.push({
      code: "shape.prop.missing",
      level: "error",
      path,
      message: "CustomGeometry requires a non-empty points array.",
    });
  }

  if (
    (node.type === "Image" || node.type === "Media") &&
    !("path" in props) &&
    !("data" in props) &&
    !("link" in props)
  ) {
    issues.push({
      code: "asset.source.missing",
      level: "warning",
      path,
      message: `${node.type} usually needs path, data, or link.`,
    });
  }
}

/** Alias for backward compatibility. */
export const validatePptxTree = validateDeck;
