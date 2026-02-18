import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";

import App from "./App";
import "./index.css";

type Theme = "light" | "dark";

declare global {
  interface Window {
    __setTheme?: (theme: Theme) => void;
  }
}

function applyTheme(theme: Theme) {
  const root = document.documentElement;
  root.classList.toggle("dark", theme === "dark");
  root.setAttribute("data-theme", theme);
}

const storedTheme = localStorage.getItem("theme") as Theme | null;
const prefersDark = window.matchMedia?.("(prefers-color-scheme: dark)").matches;
const initialTheme: Theme = storedTheme ?? (prefersDark ? "dark" : "light");
applyTheme(initialTheme);

window.__setTheme = (theme: Theme) => {
  localStorage.setItem("theme", theme);
  applyTheme(theme);
};

const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      refetchOnWindowFocus: false,
      retry: 1,
    },
    mutations: {
      retry: 0,
    },
  },
});

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <QueryClientProvider client={queryClient}>
      <App />
    </QueryClientProvider>
  </StrictMode>
);

