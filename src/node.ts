/**
 * Core node types for pptxgenjsx.
 *
 * Provides PptxNode, PptxNodePromise (async component support),
 * and type definitions used by the JSX runtime and rendering engine.
 */
import type PptxGenJS from "pptxgenjs";

// ── Primitive types ───────────────────────────────────────────────

export type PrimitiveChild = string | number | boolean | null | undefined;

export type ComponentProps = object;

// ── PptxNodePromise — wrapper for async component results ─────────

export interface PptxNodePromise {
  readonly $$pptxPromise: true;
  readonly promise: Promise<PptxNode>;
}

export function createPptxNodePromise(promise: Promise<PptxNode>): PptxNodePromise {
  return { $$pptxPromise: true, promise };
}

export function isPptxNodePromise(value: unknown): value is PptxNodePromise {
  return value != null && typeof value === "object" && (value as any).$$pptxPromise === true;
}

/** Resolve a PptxNodePromise or plain Promise<PptxNode>, otherwise return the node as-is. */
export async function resolveChild<T>(child: T): Promise<T extends PptxNodePromise ? PptxNode : T> {
  if (isPptxNodePromise(child)) {
    return (await child.promise) as any;
  }
  // Handle plain Promise<PptxNode> (e.g. from async component factories)
  if (child instanceof Promise) {
    return (await child) as any;
  }
  return child as any;
}

// ── Child types (supports both sync and async nodes) ──────────────

export type PptxChild = PptxNode | PptxNodePromise | Promise<PptxNode> | PrimitiveChild;

export type PptxChildren = PptxChild | PptxChildren[];

// ── ComponentFactory — the key async-friendly type ────────────────

/** A component factory can return a PptxNode synchronously OR a Promise. */
export type ComponentFactory<P extends ComponentProps = ComponentProps> = (
  props: P & { children?: PptxChildren },
) => PptxNode<string, P> | Promise<PptxNode<string, P>>;

/** Union type for JSX element — accepts both sync and async component results. */
export type PptxElement = PptxNode | PptxNodePromise;

// ── PptxNode class ────────────────────────────────────────────────

export class PptxNode<
  TType extends string = string,
  TProps extends ComponentProps = ComponentProps,
> {
  readonly $$pptxNode = true;
  readonly type: TType;
  readonly props: Readonly<TProps>;
  readonly children: readonly PptxChild[];

  constructor(type: TType, props: TProps & { children?: PptxChildren }) {
    const { children, ...rest } = props;
    this.type = type;
    this.props = Object.freeze(rest) as Readonly<TProps>;
    this.children = Object.freeze(flattenChildren(children));
  }

  toJSON(): Record<string, unknown> {
    return {
      type: this.type,
      props: this.props,
      children: this.children
        .filter((c): c is PptxNode | PptxNodePromise | string | number | boolean => isValidChild(c))
        .map((c) => (isPptxNode(c) ? c.toJSON() : c)),
    };
  }
}

// ── Factory helpers ───────────────────────────────────────────────

export function createNode<TType extends string, TProps extends ComponentProps>(
  type: TType,
  props: TProps & { children?: PptxChildren },
): PptxNode<TType, TProps> {
  return new PptxNode(type, props);
}

export function isPptxNode(value: unknown): value is PptxNode {
  return (
    value instanceof PptxNode || !!(value && typeof value === "object" && (value as any).$$pptxNode)
  );
}

export function isValidChild(
  value: unknown,
): value is PptxNode | PptxNodePromise | string | number | boolean {
  return (
    isPptxNode(value) ||
    isPptxNodePromise(value) ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  );
}

export function flattenChildren(children: PptxChildren | undefined): PptxChild[] {
  if (children == null || children === false) return [];
  const result: PptxChild[] = [];
  flatten(children, result);
  return result;
}

function flatten(input: PptxChildren, output: PptxChild[]): void {
  if (Array.isArray(input)) {
    for (const item of input) {
      flatten(item, output);
    }
    return;
  }
  if (input != null && input !== false) {
    output.push(input);
  }
}

// ── Re-export common pptxgenjs types for convenience ──────────────

export type PptxPresentation = PptxGenJS;
export type PptxSlide = PptxGenJS.Slide;
