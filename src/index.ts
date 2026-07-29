/**
 * pptxgenjsx — JSX runtime for building PowerPoint presentations with pptxgenjs.
 *
 * Entry point that re-exports all public API (components + context).
 *
 * @module pptxgenjsx
 */

export * from "./components.js";
export { useSlideContext, useDeckContext, useGroupContext } from "./context.js";
export type { SlideContextInfo, DeckContextInfo, GroupContextInfo } from "./context.js";
