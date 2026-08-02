/**
 * pptxgenjsx — JSX runtime for building PowerPoint presentations with pptxgenjs.
 *
 * Entry point that re-exports all public API (components + context).
 *
 * @module pptxgenjsx
 */

export type { PptxNode, PptxNodePromise, PptxChildren } from "./node.js";
export { useSlideContext, useDeckContext, useGroupContext } from "./context.js";
export type { SlideContextInfo, DeckContextInfo, GroupContextInfo } from "./context.js";
export * from "./components.js";
