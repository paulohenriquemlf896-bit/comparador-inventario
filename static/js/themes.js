/* ========================================
   THEME MANAGEMENT SYSTEM
   Dark Mode / Light Mode Toggle
   ======================================== */

class ThemeManager {
  constructor() {
    this.storageKey = 'comparador-theme';
    this.darkModeClass = 'dark-mode';
    this.dataThemeAttr = 'data-theme';
    this.init();
  }

  init() {
    // Detectar preferência do usuário ou sistema
    const savedTheme = this.getSavedTheme();
    const prefersDark = this.prefersColorScheme();
    const theme = savedTheme || (prefersDark ? 'dark' : 'light');

    this.setTheme(theme);
    this.setupSystemPreferenceListener();
  }

  getSavedTheme() {
    return localStorage.getItem(this.storageKey);
  }

  prefersColorScheme() {
    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  }

  setTheme(theme) {
    const isDark = theme === 'dark';

    if (isDark) {
      document.documentElement.setAttribute(this.dataThemeAttr, 'dark');
      document.documentElement.classList.add(this.darkModeClass);
    } else {
      document.documentElement.removeAttribute(this.dataThemeAttr);
      document.documentElement.classList.remove(this.darkModeClass);
    }

    localStorage.setItem(this.storageKey, theme);
    this.updateToggleButton(isDark);
    this.dispatchThemeChangeEvent(theme);
  }

  toggleTheme() {
    const currentTheme = this.getSavedTheme() || (this.prefersColorScheme() ? 'dark' : 'light');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    this.setTheme(newTheme);
  }

  updateToggleButton(isDark) {
    const toggle = document.getElementById('themeToggle');
    if (!toggle) return;

    toggle.setAttribute('aria-pressed', isDark);
    toggle.setAttribute('title', isDark ? 'Ativar modo claro' : 'Ativar modo escuro');

    const icon = toggle.querySelector('span');
    if (icon) {
      icon.textContent = isDark ? '☀️' : '🌙';
    }
  }

  setupSystemPreferenceListener() {
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');

    // Listener para mudanças nas preferências do sistema
    mediaQuery.addEventListener('change', (e) => {
      if (!this.getSavedTheme()) {
        this.setTheme(e.matches ? 'dark' : 'light');
      }
    });
  }

  dispatchThemeChangeEvent(theme) {
    const event = new CustomEvent('themechange', {
      detail: { theme, isDark: theme === 'dark' }
    });
    document.dispatchEvent(event);
  }

  getCurrentTheme() {
    return this.getSavedTheme() || (this.prefersColorScheme() ? 'dark' : 'light');
  }

  isDarkMode() {
    return this.getCurrentTheme() === 'dark';
  }
}

// Inicializar gerenciador de tema quando o DOM estiver pronto
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    window.themeManager = new ThemeManager();
  });
} else {
  window.themeManager = new ThemeManager();
}

// Helper functions para uso global
function toggleDarkMode() {
  if (window.themeManager) {
    window.themeManager.toggleTheme();
  }
}

function setDarkMode(isDark) {
  if (window.themeManager) {
    window.themeManager.setTheme(isDark ? 'dark' : 'light');
  }
}

function getCurrentTheme() {
  return window.themeManager ? window.themeManager.getCurrentTheme() : 'light';
}
