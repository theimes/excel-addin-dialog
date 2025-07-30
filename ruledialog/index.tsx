

/* Render application after Office initializes */
import * as React from "react";
import { createRoot } from "react-dom/client";
import Ruledialog from "./ruledialog";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("app");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady((info) => {
  console.log("Office is ready for ruledialog", info);
  
  if (root) {
    root.render(<Ruledialog />);
  }
});

// Hot module replacement for development
if ((module as any).hot) {
  (module as any).hot.accept("./ruledialog", () => {
    const NextRuledialog = require("./ruledialog").default;
    if (root) {
      root.render(<NextRuledialog />);
    }
  });
}
