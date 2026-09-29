(() => {
  "use strict";

  const storageKey = "coc7-ui-theme-v1";
  const normalizePreference = (value) => (
    value === "light" || value === "dark" ? value : "system"
  );
  const readPreference = () => {
    try {
      return normalizePreference(window.localStorage.getItem(storageKey));
    } catch (_) {
      return "system";
    }
  };

  let preference = readPreference();
  let systemTheme = null;
  let selector = null;
  try {
    if (typeof window.matchMedia === "function") {
      systemTheme = window.matchMedia("(prefers-color-scheme: dark)");
    }
  } catch (_) {
    // A manual preference still works when system-theme detection is unavailable.
  }

  const applyTheme = () => {
    const theme = preference === "system"
      ? (systemTheme && systemTheme.matches ? "dark" : "light")
      : preference;
    document.documentElement.dataset.theme = theme;
    document.documentElement.style.colorScheme = theme;
    const themeColor = document.querySelector('meta[name="theme-color"]');
    if (themeColor) themeColor.content = theme === "dark" ? "#111b1c" : "#122526";
    if (selector) selector.value = preference;
  };

  // Run before the stylesheet loads so a saved preference also applies on reload.
  applyTheme();

  if (systemTheme) {
    if (typeof systemTheme.addEventListener === "function") {
      systemTheme.addEventListener("change", applyTheme);
    } else if (typeof systemTheme.addListener === "function") {
      systemTheme.addListener(applyTheme);
    }
  }

  window.addEventListener("storage", (event) => {
    if (event.key !== storageKey && event.key !== null) return;
    preference = event.key === null ? readPreference() : normalizePreference(event.newValue);
    applyTheme();
  });

  const bindSelector = () => {
    selector = document.getElementById("theme-select");
    if (!selector) return;
    selector.value = preference;
    selector.addEventListener("change", () => {
      preference = normalizePreference(selector.value);
      try {
        window.localStorage.setItem(storageKey, preference);
      } catch (_) {
        // Keep the current-page choice usable when browser storage is disabled.
      }
      applyTheme();
    });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindSelector, { once: true });
  } else {
    bindSelector();
  }
})();
