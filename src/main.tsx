import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./ui/app";

Office.onReady(() => {
  createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
});
