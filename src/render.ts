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
    await processDeckChildren(node, { pptx });
    return pptx;
  }

  if (node.type === "Slide") {
    await processSlide(node, { pptx });
    return pptx;
  }

  throw new Error(`createPptx expected a Deck or Slide root, got ${node.type}.`);
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

async function processDeckChildren(node: PptxNode, ctx: { pptx: PptxPresentation }): Promise<void> {
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

async function processSection(node: PptxNode, ctx: { pptx: PptxPresentation }): Promise<void> {
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
  ctx: { pptx: PptxPresentation },
  sectionTitle?: string,
): Promise<void> {
  const props = node.props as any;
  const slide = ctx.pptx.addSlide({
    masterName: props.masterName,
    sectionTitle: props.sectionTitle ?? sectionTitle,
  });

  setIfDefined(slide, "background", props.background);
  setIfDefined(slide, "color", props.color);
  setIfDefined(slide, "hidden", props.hidden);
  setIfDefined(slide, "slideNumber", props.slideNumber);

  for (const rawChild of node.children) {
    const child = await resolveChild(rawChild);
    if (!isPptxNode(child)) continue;
    await renderSlideElement(child, { ...ctx, slide });
  }
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

// ── Individual element renderers ──────────────────────────────────

function renderText(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
  const textRuns = collectTextRuns(node) ?? props.text ?? getTextContent(node);
  slide.addText(textRuns, mergeOptions(props.options, props));
}

function renderShape(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
  slide.addShape(props.shape, mergeOptions(props.options, props, ["shape"]));
}

function renderLine(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
  slide.addShape("line", mergeOptions(props.options, props));
}

function renderLineBetween(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
  const x = Math.min(props.x1, props.x2);
  const y = Math.min(props.y1, props.y2);
  const w = Math.abs(props.x2 - props.x1);
  const h = Math.abs(props.y2 - props.y1);

  slide.addShape("line", {
    ...mergeOptions(props.options, props, ["x1", "y1", "x2", "y2"]),
    x,
    y,
    w,
    h,
    flipH: props.x2 < props.x1,
    flipV: props.y2 < props.y1,
  });
}

function renderImage(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
  slide.addImage(mergeOptions(props.options, props));
}

function renderMedia(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
  slide.addMedia(mergeOptions(props.options, props) as any);
}

function renderChart(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
  slide.addChart(
    props.type,
    deepClone(props.data),
    mergeOptions(props.options, props, ["type", "data"]),
  );
}

function renderTable(node: PptxNode, slide: PptxSlide): void {
  const props = node.props as any;
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
    for (const key of ["x1", "y1", "x2", "y2"]) {
      if (typeof props[key] !== "number" || Number.isNaN(props[key])) {
        issues.push({
          code: "line.endpoint.invalid",
          level: "error",
          path,
          message: `LineBetween requires numeric ${key} in PPT inches.`,
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
