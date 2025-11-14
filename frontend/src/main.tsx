import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "bootstrap/dist/css/bootstrap.min.css";
import "./index.css";

// Modo oscuro/claro con Bootstrap 5 (data-bs-theme) + persistencia
type Theme = "light" | "dark";
declare global {
  interface Window {
    __setTheme?: (t: Theme) => void;
  }
}
function applyTheme(t: Theme) {
  document.documentElement.setAttribute("data-bs-theme", t);
  if (t === "dark") {
    document.body.classList.add("theme-dark");
  } else {
    document.body.classList.remove("theme-dark");
  }
}
const stored = (localStorage.getItem("theme") as Theme | null);
const preferDark = window.matchMedia?.("(prefers-color-scheme: dark)").matches;
const initialTheme: Theme = stored ?? (preferDark ? "dark" : "light");
applyTheme(initialTheme);
window.__setTheme = (t: Theme) => {
  localStorage.setItem("theme", t);
  applyTheme(t);
};
document.body.classList.add("adminator-skin");

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App />
  </StrictMode>
);
