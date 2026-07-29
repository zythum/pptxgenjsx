/**
 * JSX runtime for pptxgenjsx.
 *
 * Provides jsx/jsxs/jsxDEV/Fragment for TypeScript's `react-jsx` transform.
 * Async component factories are wrapped in PptxNodePromise for lazy resolution
 * during tree traversal (in render.ts).
 */
import {
  type ComponentFactory,
  type ComponentProps,
  type PptxChild,
  type PptxChildren,
  type PptxNode,
  type PptxElement,
  createNode,
  flattenChildren,
  createPptxNodePromise,
} from "./node.js";

// ── JSX factory ───────────────────────────────────────────────────

/**
 * JSX factory — called by TypeScript's `react-jsx` transform.
 *
 * If the Component returns a Promise (async component), it is wrapped in
 * a PptxNodePromise and resolved lazily during `createPptx` traversal.
 */
export function jsx<P extends ComponentProps>(
  Component: ComponentFactory<P> | string,
  props: P & { children?: PptxChildren },
  _key?: string,
): PptxElement {
  const { children, ...rest } = props ?? ({} as any);
  const flatChildren = flattenChildren(children);

  if (typeof Component === "string") {
    // Intrinsic element (not used in this runtime, but kept for completeness)
    return createNode(Component, {
      ...rest,
      children: flatChildren,
    } as any);
  }

  const result = Component({
    ...rest,
    children: flatChildren,
  } as P & { children?: PptxChildren });

  // Wrap async component results
  if (result instanceof Promise) {
    return createPptxNodePromise(result);
  }

  return result;
}

/** Alias for jsx (used when there are multiple children). */
export const jsxs: typeof jsx = jsx;

/** Alias for dev-mode (same as jsx in this runtime). */
export const jsxDEV: typeof jsx = jsx;

// ── Fragment ──────────────────────────────────────────────────────

export function Fragment(props: { children?: PptxChildren }): PptxNode<"Fragment", {}> {
  return createNode("Fragment", {
    children: flattenChildren(props.children),
  });
}

// ── JSX namespace for TypeScript ──────────────────────────────────

export namespace JSX {
  /**
   * Accept both sync (PptxNode) and async (Promise<PptxNode>) component results.
   * TypeScript checks the component factory's declared return type against JSX.Element.
   * Async results are wrapped in PptxNodePromise by the jsx() factory at runtime.
   */
  export type Element = PptxNode | Promise<PptxNode>;
  export interface ElementChildrenAttribute {
    children: unknown;
  }
  export interface IntrinsicElements {}
}

// ── Additional factory aliases (h / pptxElement / createElement) ──

/**
 * Hyperscript-style factory.
 * Prefer `pptxElement` to avoid shadowing PptxGenJS height props.
 */
export function h<P extends ComponentProps>(
  Component: ComponentFactory<P>,
  props: P | null,
  ...children: PptxChild[]
): PptxElement {
  return jsx(Component, { ...(props ?? ({} as P)), children } as any);
}

/** Alias that avoids shadowing the `h` (height) prop. */
export const pptxElement = h;

/** React-compatible alias. */
export const createElement = h;
