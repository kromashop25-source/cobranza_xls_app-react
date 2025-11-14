import { useEffect, useState } from "react";

type Theme = "light" | "dark";

export default function ThemeToggle() {
  const getTheme = (): Theme =>
    (document.documentElement.getAttribute("data-bs-theme") as Theme) === "dark"
      ? "dark"
      : "light";
  const [theme, setTheme] = useState<Theme>(getTheme());

  useEffect(() => {
    const obs = new MutationObserver(() => setTheme(getTheme()));
    obs.observe(document.documentElement, { attributes: true, attributeFilter: ["data-bs-theme"] });
    return () => obs.disconnect();
  }, []);

  const toggle = () => {
    const next: Theme = theme === "dark" ? "light" : "dark";
    window.__setTheme?.(next);
    setTheme(next);
  };

  return (
    <button
      type="button"
      className={`btn btn-sm ${theme === "dark" ? "btn-outline-light" : "btn-outline-secondary"}`}
      onClick={toggle}
      title={theme === "dark" ? "Cambiar a modo claro" : "Cambiar a modo oscuro"}
    >
      {theme === "dark" ? "Modo claro" : "Modo oscuro"}
    </button>
  );
}