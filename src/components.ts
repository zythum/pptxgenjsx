/**
 * Component factory definitions for pptxgenjsx.
 *
 * Each component factory creates a PptxNode with a specific type tag.
 * All positioning values are in inches (canvas: 13.333 x 7.5 for WIDE).
 */

import type PptxGenJSType from "pptxgenjs";

import {
  type ComponentFactory,
  type PptxChildren,
  type PptxNode,
  type PptxNodePromise,
  createNode,
} from "./node.js";

// Re-export core types
export type { PptxNode, PptxNodePromise };

// ════════════════════════════════════════════════════════════════════
// COMPONENT PROPS TYPE DEFINITIONS
// ════════════════════════════════════════════════════════════════════

// ── Presentation ──────────────────────────────────────────────────

export type DeckProps = {
  /** Presentation title metadata. */
  title?: string;
  /** Presentation author metadata. */
  author?: string;
  /** Presentation company metadata. */
  company?: string;
  /** Presentation subject metadata. */
  subject?: string;
  /** Presentation revision metadata. Must be a whole-number string for PowerPoint compatibility. */
  revision?: string;
  /** Enable right-to-left mode for the whole presentation. */
  rtlMode?: boolean;
  /** Default theme fonts. Maps to `pptx.theme`. */
  theme?: PptxGenJSType.ThemeProps;
  /** Built-in layout name or custom layout object. Use `layouts` for additional custom layouts. */
  layout?: string | PptxGenJSType.PresLayout;
  /** Additional custom layouts passed to `pptx.defineLayout`. */
  layouts?: PptxGenJSType.PresLayout[];
  /** Master slides passed to `pptx.defineSlideMaster`. */
  masters?: PptxGenJSType.SlideMasterProps[];
  /** Presentation sections passed to `pptx.addSection`. */
  sections?: PptxGenJSType.SectionProps[];
  children?: PptxChildren;
};

// ── Structure ─────────────────────────────────────────────────────

export type SlideProps = Partial<
  Pick<PptxGenJSType.PresSlide, "background" | "color" | "hidden" | "slideNumber">
> &
  PptxGenJSType.AddSlideProps & {
    children?: PptxChildren;
  };

export type LayoutProps = PptxGenJSType.PresLayout;

export type SectionProps = PptxGenJSType.SectionProps & {
  children?: PptxChildren;
};

export type MasterProps = PptxGenJSType.SlideMasterProps & {
  children?: PptxChildren;
};

export type PlaceholderProps = {
  /** Placeholder options used inside `defineSlideMaster({ objects })`. */
  options: PptxGenJSType.PlaceholderProps;
  /** Placeholder text shown until edited. Children strings are joined when `text` is omitted. */
  text?: string;
  children?: PptxChildren;
};

// ── Text ──────────────────────────────────────────────────────────

export type TextRunProps = PptxGenJSType.TextProps;

export type TextProps = Omit<PptxGenJSType.TextPropsOptions, "children"> & {
  /** Text string passed as first argument to `slide.addText`. Children strings are joined when `text` is omitted. */
  text?: string | PptxGenJSType.TextProps[];
  /** Rich text runs passed as first argument to `slide.addText`. Same shape as PptxGenJS `TextProps[]`. */
  runs?: PptxGenJSType.TextProps[];
  /** Optional nested `<TextRun />` nodes or plain strings. */
  children?: PptxChildren;
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.TextPropsOptions;
};

export type NotesProps = {
  /** Speaker notes string. Children strings are joined when `text` is omitted. */
  text?: string;
  children?: PptxChildren;
};

// ── Shapes ────────────────────────────────────────────────────────

export type ShapeProps = Omit<PptxGenJSType.ShapeProps, "children"> & {
  /** Shape type passed as first argument to `slide.addShape`. Example: `rect`, `ellipse`, `roundRect`, `line`. */
  shape: PptxGenJSType.SHAPE_NAME;
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.ShapeProps;
};

export type LineProps = Omit<ShapeProps, "shape">;

export type LineBetweenProps = Omit<LineProps, "x" | "y" | "w" | "h" | "flipH" | "flipV"> & {
  /** Start x coordinate in PPT inches. */
  x1: number;
  /** Start y coordinate in PPT inches. */
  y1: number;
  /** End x coordinate in PPT inches. */
  x2: number;
  /** End y coordinate in PPT inches. */
  y2: number;
};

export type ShapeOptionsProps = Omit<PptxGenJSType.ShapeProps, "children"> & {
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.ShapeProps;
};

export type RectProps = ShapeOptionsProps;
export type EllipseProps = ShapeOptionsProps;
export type TriangleProps = ShapeOptionsProps;
export type DiamondProps = ShapeOptionsProps;
export type PentagonProps = ShapeOptionsProps;
export type HexagonProps = ShapeOptionsProps;
export type StarProps = ShapeOptionsProps;
export type ArrowProps = ShapeOptionsProps;

export type RoundRectProps = ShapeOptionsProps & {
  /** Rounded rectangle radius. Valid only for `roundRect`; range 0.0 to 1.0. */
  rectRadius?: number;
};

export type ArcProps = ShapeOptionsProps & {
  /** Arc angle range. Valid for `arc`, `pie`, and `blockArc`; range [0-359, 0-359]. */
  angleRange?: [number, number];
};

export type BlockArcProps = ArcProps & {
  /** Block arc thickness ratio. Valid only for `blockArc`; range 0.0 to 1.0. */
  arcThicknessRatio?: number;
};

export type PieShapeProps = ArcProps;

export type CustomGeometryPoint =
  | { x: PptxGenJSType.Coord; y: PptxGenJSType.Coord; moveTo?: boolean }
  | {
      x: PptxGenJSType.Coord;
      y: PptxGenJSType.Coord;
      curve: {
        type: "arc";
        hR: PptxGenJSType.Coord;
        wR: PptxGenJSType.Coord;
        stAng: number;
        swAng: number;
      };
    }
  | {
      x: PptxGenJSType.Coord;
      y: PptxGenJSType.Coord;
      curve: {
        type: "cubic";
        x1: PptxGenJSType.Coord;
        y1: PptxGenJSType.Coord;
        x2: PptxGenJSType.Coord;
        y2: PptxGenJSType.Coord;
      };
    }
  | {
      x: PptxGenJSType.Coord;
      y: PptxGenJSType.Coord;
      curve: { type: "quadratic"; x1: PptxGenJSType.Coord; y1: PptxGenJSType.Coord };
    }
  | { close: true };

export type CustomGeometryProps = ShapeOptionsProps & {
  /**
   * Custom geometry path points passed to PptxGenJS `custGeom`.
   * Convert SVG path commands before passing them here:
   * `M` -> `{ x, y, moveTo:true }`, `L` -> `{ x, y }`, `C` -> cubic curve,
   * `Q` -> quadratic curve, `Z` -> `{ close:true }`.
   */
  points: CustomGeometryPoint[];
};

// ── Media ─────────────────────────────────────────────────────────

export type ImageProps = PptxGenJSType.ImageProps & {
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.ImageProps;
};

export type MediaProps = PptxGenJSType.MediaProps & {
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.MediaProps;
};

// ── Charts ────────────────────────────────────────────────────────

export type ChartProps = Omit<PptxGenJSType.IChartOpts, "children"> & {
  /** Chart type or multi-chart descriptor. */
  type: PptxGenJSType.CHART_NAME | PptxGenJSType.IChartMulti[];
  /** Chart series data. */
  data: PptxGenJSType.OptsChartData[];
  /** Base chart options merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.IChartOpts;
};

export type TypedChartProps = Omit<ChartProps, "type">;

export type AreaChartProps = TypedChartProps;
export type BarChartProps = TypedChartProps;
export type Bar3DChartProps = TypedChartProps;
export type BubbleChartProps = TypedChartProps;
export type DoughnutChartProps = TypedChartProps;
export type LineChartProps = TypedChartProps;
export type PieChartProps = TypedChartProps;
export type RadarChartProps = TypedChartProps;
export type ScatterChartProps = TypedChartProps;

// ── Table ─────────────────────────────────────────────────────────

export type TableCellProps = PptxGenJSType.TableCellProps & {
  /** Cell text. Children strings are joined when `text` is omitted. */
  text?: string | PptxGenJSType.TableCell[];
  children?: PptxChildren;
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.TableCellProps;
};

export type TableRowProps = {
  /** Optional direct row data. When omitted, nested `<TableCell />` nodes are collected. */
  cells?: PptxGenJSType.TableRow;
  children?: PptxChildren;
};

export type TableProps = Omit<PptxGenJSType.TableProps, "children"> & {
  /** Table rows passed as first argument to `slide.addTable`. */
  rows?: PptxGenJSType.TableRow[];
  children?: PptxChildren;
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.TableProps;
};

export type TableToSlidesProps = PptxGenJSType.TableToSlidesProps & {
  /** HTML table element id passed to `pptx.tableToSlides`. */
  eleId: string;
  /** Base options object merged with top-level option props. Top-level props win. */
  options?: PptxGenJSType.TableToSlidesProps;
};

// ── Escape hatch ──────────────────────────────────────────────────

export type RawProps = {
  /** Escape hatch for unsupported or newly added PptxGenJS APIs. */
  render: (context: RenderContext) => void | Promise<void>;
};

export type RenderContext = {
  pptx: PptxGenJSType;
  slide?: PptxGenJSType.PresSlide;
  node: PptxNode<"Raw">;
};

// ════════════════════════════════════════════════════════════════════
// COMPONENT FACTORIES
// ════════════════════════════════════════════════════════════════════

// ── Generic factory helpers ───────────────────────────────────────

function component<T extends string>(type: T): ComponentFactory {
  return (props) => createNode(type, props as any);
}

function shapeComponent<T extends string>(type: T, shape: string): ComponentFactory {
  return (props) => createNode(type, { ...props, shape } as any);
}

function chartComponent<T extends string>(type: T, chartType: string): ComponentFactory {
  return (props) => createNode(type, { ...props, type: chartType } as any);
}

// ── Deck / Presentation ──────────────────────────────────────────

/** Root presentation. Maps to `new PptxGenJS()`. */
export const Deck = component("Deck") as ComponentFactory<DeckProps>;
/** Alias for Deck. */
export const Presentation = Deck as ComponentFactory<DeckProps>;

// ── Slide ─────────────────────────────────────────────────────────

/** Create a slide. Maps to `pptx.addSlide`. */
export const Slide = component("Slide") as ComponentFactory<SlideProps>;

// ── Layout / Section / Master / Placeholder ───────────────────────

/** Define an additional custom layout. Maps to `pptx.defineLayout`. */
export const Layout = component("Layout") as ComponentFactory<LayoutProps>;
/** Group nested slides into a PowerPoint section. Maps to `pptx.addSection`. */
export const Section = component("Section") as ComponentFactory<SectionProps>;
/** Define a reusable slide master. Maps to `pptx.defineSlideMaster`. */
export const Master = component("Master") as ComponentFactory<MasterProps>;
/** Add a placeholder inside Master. Maps to `SlideMasterProps.objects[].placeholder`. */
export const Placeholder = component("Placeholder") as ComponentFactory<PlaceholderProps>;

// ── Text ──────────────────────────────────────────────────────────

/** Text box, rich text container, or text inside a shape. Maps to `slide.addText`. */
export const Text = component("Text") as ComponentFactory<TextProps>;
/** Per-run formatting inside Text. Maps to `PptxGenJS.TextProps[]` item. */
export const TextRun = component("TextRun") as ComponentFactory<TextRunProps>;
/** Speaker notes. Maps to `slide.addNotes`. */
export const Notes = component("Notes") as ComponentFactory<NotesProps>;

// ── Shapes ────────────────────────────────────────────────────────

/** Any PptxGenJS shape not exposed as a dedicated component. Maps to `slide.addShape(shape, options)`. */
export const Shape = component("Shape") as ComponentFactory<ShapeProps>;

/** Pure rectangle. Maps to `slide.addShape('rect')`. */
export const Rect = shapeComponent("Rect", "rect") as ComponentFactory<RectProps>;
/** Rounded rectangle; supports rectRadius. Maps to `slide.addShape('roundRect')`. */
export const RoundRect = shapeComponent(
  "RoundRect",
  "roundRect",
) as ComponentFactory<RoundRectProps>;
/** Ellipse or oval. Maps to `slide.addShape('ellipse')`. */
export const Ellipse = shapeComponent("Ellipse", "ellipse") as ComponentFactory<EllipseProps>;
/** Alias for Ellipse. */
export const Oval = Ellipse;
/** Triangle shape. Maps to `slide.addShape('triangle')`. */
export const Triangle = shapeComponent("Triangle", "triangle") as ComponentFactory<TriangleProps>;
/** Right triangle shape. Maps to `slide.addShape('rtTriangle')`. */
export const RightTriangle = shapeComponent(
  "RightTriangle",
  "rtTriangle",
) as ComponentFactory<TriangleProps>;
/** Diamond shape. Maps to `slide.addShape('diamond')`. */
export const Diamond = shapeComponent("Diamond", "diamond") as ComponentFactory<DiamondProps>;
/** Pentagon shape. Maps to `slide.addShape('pentagon')`. */
export const Pentagon = shapeComponent(
  "Pentagon",
  "pentagon",
) as ComponentFactory<ShapeOptionsProps>;
/** Hexagon shape. Maps to `slide.addShape('hexagon')`. */
export const Hexagon = shapeComponent("Hexagon", "hexagon") as ComponentFactory<HexagonProps>;
/** Five-point star shape. Maps to `slide.addShape('star5')`. */
export const Star = shapeComponent("Star", "star5") as ComponentFactory<StarProps>;
/** Four-point star shape. Maps to `slide.addShape('star4')`. */
export const Star4 = shapeComponent("Star4", "star4") as ComponentFactory<StarProps>;
/** Alias for Star (five-point). */
export const Star5 = Star;
/** Six-point star shape. Maps to `slide.addShape('star6')`. */
export const Star6 = shapeComponent("Star6", "star6") as ComponentFactory<StarProps>;
/** Eight-point star shape. Maps to `slide.addShape('star8')`. */
export const Star8 = shapeComponent("Star8", "star8") as ComponentFactory<StarProps>;
/** Ten-point star shape. Maps to `slide.addShape('star10')`. */
export const Star10 = shapeComponent("Star10", "star10") as ComponentFactory<StarProps>;
/** Line or arrow line. Maps to `slide.addShape('line')`. */
export const Line = component("Line") as ComponentFactory<LineProps>;
/** Line between two endpoints. Maps to `slide.addShape('line')` with normalized x/y/w/h and flipH/flipV. */
export const LineBetween = component("LineBetween") as ComponentFactory<LineBetweenProps>;
/** Arc shape; supports angleRange. Maps to `slide.addShape('arc')`. */
export const Arc = shapeComponent("Arc", "arc") as ComponentFactory<ArcProps>;
/** Thick arc; supports angleRange and arcThicknessRatio. Maps to `slide.addShape('blockArc')`. */
export const BlockArc = shapeComponent("BlockArc", "blockArc") as ComponentFactory<BlockArcProps>;
/** Pie-slice shape; not a data chart. Maps to `slide.addShape('pie')`. */
export const PieShape = shapeComponent("PieShape", "pie") as ComponentFactory<PieShapeProps>;
/** Editable custom path geometry. Maps to `slide.addShape('custGeom')`. */
export const CustomGeometry = shapeComponent(
  "CustomGeometry",
  "custGeom",
) as ComponentFactory<CustomGeometryProps>;
/** Left-pointing arrow. Maps to `slide.addShape('leftArrow')`. */
export const LeftArrow = shapeComponent("LeftArrow", "leftArrow") as ComponentFactory<ArrowProps>;
/** Right-pointing arrow. Maps to `slide.addShape('rightArrow')`. */
export const RightArrow = shapeComponent(
  "RightArrow",
  "rightArrow",
) as ComponentFactory<ArrowProps>;
/** Up-pointing arrow. Maps to `slide.addShape('upArrow')`. */
export const UpArrow = shapeComponent("UpArrow", "upArrow") as ComponentFactory<ArrowProps>;
/** Down-pointing arrow. Maps to `slide.addShape('downArrow')`. */
export const DownArrow = shapeComponent("DownArrow", "downArrow") as ComponentFactory<ArrowProps>;
/** Left-right bidirectional arrow. Maps to `slide.addShape('leftRightArrow')`. */
export const LeftRightArrow = shapeComponent(
  "LeftRightArrow",
  "leftRightArrow",
) as ComponentFactory<ArrowProps>;
/** Up-down bidirectional arrow. Maps to `slide.addShape('upDownArrow')`. */
export const UpDownArrow = shapeComponent(
  "UpDownArrow",
  "upDownArrow",
) as ComponentFactory<ArrowProps>;
/** Chevron / angle bracket shape. Maps to `slide.addShape('chevron')`. */
export const Chevron = shapeComponent("Chevron", "chevron") as ComponentFactory<ArrowProps>;
/** Cloud shape. Maps to `slide.addShape('cloud')`. */
export const Cloud = shapeComponent("Cloud", "cloud") as ComponentFactory<ShapeOptionsProps>;
/** Heart shape. Maps to `slide.addShape('heart')`. */
export const Heart = shapeComponent("Heart", "heart") as ComponentFactory<ShapeOptionsProps>;
/** Donut / ring shape. Maps to `slide.addShape('donut')`. */
export const Donut = shapeComponent("Donut", "donut") as ComponentFactory<ShapeOptionsProps>;
/** Plus / cross shape. Maps to `slide.addShape('plus')`. */
export const Plus = shapeComponent("Plus", "plus") as ComponentFactory<ShapeOptionsProps>;

// ── Media ─────────────────────────────────────────────────────────

/** Image by path or data URI/base64. Maps to `slide.addImage`. */
export const Image = component("Image") as ComponentFactory<ImageProps>;
/** Audio, video, or online media. Maps to `slide.addMedia`. */
export const Media = component("Media") as ComponentFactory<MediaProps>;

// ── Charts ────────────────────────────────────────────────────────

/** Multi-chart or any raw chart type usage. Maps to `slide.addChart(type, data, options)`. */
export const Chart = component("Chart") as ComponentFactory<ChartProps>;
/** Area chart. Maps to `slide.addChart('area')`. */
export const AreaChart = chartComponent("AreaChart", "area") as ComponentFactory<AreaChartProps>;
/** Bar chart. Maps to `slide.addChart('bar')`. */
export const BarChart = chartComponent("BarChart", "bar") as ComponentFactory<BarChartProps>;
/** 3D bar chart. Maps to `slide.addChart('bar3D')`. */
export const Bar3DChart = chartComponent(
  "Bar3DChart",
  "bar3D",
) as ComponentFactory<Bar3DChartProps>;
/** Bubble chart. Maps to `slide.addChart('bubble')`. */
export const BubbleChart = chartComponent(
  "BubbleChart",
  "bubble",
) as ComponentFactory<BubbleChartProps>;
/** Doughnut chart. Maps to `slide.addChart('doughnut')`. */
export const DoughnutChart = chartComponent(
  "DoughnutChart",
  "doughnut",
) as ComponentFactory<DoughnutChartProps>;
/** Line chart. Maps to `slide.addChart('line')`. */
export const LineChart = chartComponent("LineChart", "line") as ComponentFactory<LineChartProps>;
/** Pie chart. Maps to `slide.addChart('pie')`. */
export const PieChart = chartComponent("PieChart", "pie") as ComponentFactory<PieChartProps>;
/** Radar chart. Maps to `slide.addChart('radar')`. */
export const RadarChart = chartComponent(
  "RadarChart",
  "radar",
) as ComponentFactory<RadarChartProps>;
/** Scatter chart. Maps to `slide.addChart('scatter')`. */
export const ScatterChart = chartComponent(
  "ScatterChart",
  "scatter",
) as ComponentFactory<ScatterChartProps>;

// ── Table ─────────────────────────────────────────────────────────

/** PowerPoint table. Maps to `slide.addTable`. */
export const Table = component("Table") as ComponentFactory<TableProps>;
/** Declarative row inside Table. Maps to `PptxGenJS.TableRow`. */
export const TableRow = component("TableRow") as ComponentFactory<TableRowProps>;
/** Declarative cell inside TableRow. Maps to `PptxGenJS.TableCell`. */
export const TableCell = component("TableCell") as ComponentFactory<TableCellProps>;
/** Convert an HTML table in browser/runtime DOM. Maps to `pptx.tableToSlides`. */
export const TableToSlides = component("TableToSlides") as ComponentFactory<TableToSlidesProps>;

// ── Escape hatch ──────────────────────────────────────────────────

/** Newest or unsupported PptxGenJS APIs. Maps to custom render callback. */
export const Raw = component("Raw") as ComponentFactory<RawProps>;

// ════════════════════════════════════════════════════════════════════
// DEFAULT EXPORT — all component factories as a single object
// ════════════════════════════════════════════════════════════════════

/** All component factories as a single object. Import as default:
 * ```ts
 * import P from "@zythum02/pptxgenjsx";
 * P.Deck, P.Slide, P.Rect, ...
 * ```
 */
const _default = {
  Deck,
  Presentation,
  Slide,
  Layout,
  Section,
  Master,
  Placeholder,
  Text,
  TextRun,
  Notes,
  Shape,
  Rect,
  RoundRect,
  Ellipse,
  Oval,
  Triangle,
  RightTriangle,
  Diamond,
  Pentagon,
  Hexagon,
  Star,
  Star4,
  Star5,
  Star6,
  Star8,
  Star10,
  Line,
  LineBetween,
  Arc,
  BlockArc,
  PieShape,
  CustomGeometry,
  LeftArrow,
  RightArrow,
  UpArrow,
  DownArrow,
  LeftRightArrow,
  UpDownArrow,
  Chevron,
  Cloud,
  Heart,
  Donut,
  Plus,
  Image,
  Media,
  Chart,
  AreaChart,
  BarChart,
  Bar3DChart,
  BubbleChart,
  DoughnutChart,
  LineChart,
  PieChart,
  RadarChart,
  ScatterChart,
  Table,
  TableRow,
  TableCell,
  TableToSlides,
  Raw,
};

export default _default;
