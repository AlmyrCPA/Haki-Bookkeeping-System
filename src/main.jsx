import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

const root = document.getElementById("root");
// Full-viewport so the app's flex sidebar/main layout fills the page.
document.documentElement.style.height = "100%";
document.body.style.height = "100%";
document.body.style.margin = "0";

ReactDOM.createRoot(root).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
