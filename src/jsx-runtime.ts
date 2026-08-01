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

// ── JSX Key ───────────────────────────────────────────────────

export type Key = string | number | bigint;

// ── JSX factory ───────────────────────────────────────────────────

/**
 * JSX factory — called by TypeScript's `react-jsx` transform.
 *
 * If the Component returns a Promise (async component), it is wrapped in
 * a PptxNodePromise and resolved lazily during `createPptx` traversal.
 */
export function jsx<P extends ComponentProps>(
  Component: ComponentFactory<P> | string,
  props: P & { key?: Key; children?: PptxChildren },
  _key?: Key | null,
): PptxElement {
  const { children, key: _unusedKey, ...rest } = props ?? ({} as any);
  const flatChildren = flattenChildren(children);

  if (typeof Component === "string") {
    // Intrinsic element (not used in this runtime, but kept for completeness)
    return createNode(Component, {
      ...rest,
      children: flatChildren,
    } as any);
  }

  // ── Deferred execution via synchronous thunk ──────────────────────────
  //
  // Defer ALL non-string component factory calls so they execute during
  // tree traversal (in render.ts) rather than during JSX construction.
  //
  // Why?  Components like <Slide><Title /></Slide> evaluate children first
  // (jsx(Title, {}) runs before jsx(Slide, ...)).  If Title's factory calls
  // useSlideContext(), useDeckContext(), or useGroupContext(), those contexts
  // DON'T exist yet during JSX construction — they are set up in render.ts
  // inside withSlideContext() / withDeckContext() / withGroupContext().
  //
  // By wrapping every factory call in a PptxNodePromise **thunk** (a zero-arg
  // function), the factory is executed lazily when resolveChild() processes
  // the promise.  The thunk is called **synchronously** from resolveChild(),
  // so the factory inherits whatever rendering context (slide / deck / group)
  // is active at the call site.
  //
  // This is critical: using Promise.resolve().then() would break
  // AsyncLocalStorage propagation because the promise chain is created
  // during JSX construction (outside any rendering context).  The thunk
  // approach keeps the factory call synchronous from the perspective of
  // the rendering engine's event loop.
  // ───────────────────────────────────────────────────────────────────────
  return createPptxNodePromise(() =>
    Component({
      ...rest,
      children: flatChildren,
    } as P & { children?: PptxChildren }),
  );
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
  /**
   * React compatibility: allow `key` on every JSX element so that
   * React-style code (e.g. `key` inside `.map()` loops) type-checks.
   * The key is stripped at runtime and has no effect on rendering.
   */
  export interface IntrinsicAttributes {
    key?: Key | null;
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
  props: (P & { key?: Key | null }) | null,
  ...children: PptxChild[]
): PptxElement {
  return jsx(Component, { ...props, children } as any, props?.key);
}

/** Alias that avoids shadowing the `h` (height) prop. */
export const pptxElement = h;

/** React-compatible alias. */
export const createElement = h;
