import { useEffect, useState } from "react";

export type Theme = "light" | "dark";

const STORAGE_KEY = "checkin-quality-theme";
const DARK_MODE_QUERY = "(prefers-color-scheme: dark)";

function storedTheme(): Theme | null {
  const value = window.localStorage.getItem(STORAGE_KEY);
  return value === "light" || value === "dark" ? value : null;
}

function systemTheme(): Theme {
  return window.matchMedia(DARK_MODE_QUERY).matches ? "dark" : "light";
}

export function initialTheme(): Theme {
  return storedTheme() || systemTheme();
}

export function applyTheme(theme: Theme) {
  document.documentElement.dataset.theme = theme;
  document.documentElement.style.colorScheme = theme;
}

export function useTheme() {
  const [theme, setTheme] = useState<Theme>(initialTheme);
  const [usesSystemTheme, setUsesSystemTheme] = useState(() => storedTheme() === null);

  useEffect(() => {
    applyTheme(theme);
  }, [theme]);

  useEffect(() => {
    if (!usesSystemTheme) return;
    const mediaQuery = window.matchMedia(DARK_MODE_QUERY);
    const handleChange = (event: MediaQueryListEvent) => {
      setTheme(event.matches ? "dark" : "light");
    };
    mediaQuery.addEventListener("change", handleChange);
    return () => mediaQuery.removeEventListener("change", handleChange);
  }, [usesSystemTheme]);

  function toggleTheme() {
    const nextTheme = theme === "dark" ? "light" : "dark";
    window.localStorage.setItem(STORAGE_KEY, nextTheme);
    setUsesSystemTheme(false);
    setTheme(nextTheme);
  }

  return { theme, usesSystemTheme, toggleTheme };
}
