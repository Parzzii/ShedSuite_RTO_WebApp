(() => {
  const STORAGE_KEY = 'shedsuite-theme';
  const root = document.documentElement;
  const allowed = new Set(['light', 'dark', 'cyberpunk']);

  function storedTheme() {
    try {
      const value = localStorage.getItem(STORAGE_KEY) || 'light';
      return allowed.has(value) ? value : 'light';
    } catch (_) {
      return 'light';
    }
  }

  function saveTheme(theme) {
    try { localStorage.setItem(STORAGE_KEY, theme); } catch (_) {}
  }

  function currentTheme() {
    const theme = root.dataset.theme || storedTheme();
    return allowed.has(theme) ? theme : 'light';
  }

  function updateButton() {
    const btn = document.getElementById('themeToggle');
    if (!btn) return;
    const theme = currentTheme();
    if (theme === 'light') {
      btn.innerHTML = '<span aria-hidden="true">☾</span><span>Dark mode</span>';
      btn.setAttribute('aria-label', 'Switch to dark mode');
      btn.title = 'Switch to dark mode';
    } else if (theme === 'dark') {
      btn.innerHTML = '<span aria-hidden="true">☀</span><span>Light mode</span>';
      btn.setAttribute('aria-label', 'Switch to light mode');
      btn.title = 'Switch to light mode';
    } else {
      // Keep the secret theme out of the normal theme menu.
      btn.innerHTML = '<span aria-hidden="true">◐</span><span>Dark mode</span>';
      btn.setAttribute('aria-label', 'Return to normal dark mode');
      btn.title = 'Return to normal dark mode';
    }
  }

  function setTheme(theme, { toast = false } = {}) {
    if (!allowed.has(theme)) theme = 'light';
    if (theme === 'light') delete root.dataset.theme;
    else root.dataset.theme = theme;
    saveTheme(theme);
    updateButton();
    if (toast) showCyberToast(theme === 'cyberpunk');
  }

  function showCyberToast(enabled) {
    let toast = document.getElementById('cyberToast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'cyberToast';
      toast.className = 'cyber-toast';
      document.body.appendChild(toast);
    }
    toast.textContent = enabled ? 'NEON LINK // ONLINE' : 'NEON LINK // OFFLINE';
    toast.classList.remove('show');
    void toast.offsetWidth;
    toast.classList.add('show');
    window.clearTimeout(showCyberToast.timer);
    showCyberToast.timer = window.setTimeout(() => toast.classList.remove('show'), 1800);
  }

  function toggleNormalTheme() {
    const theme = currentTheme();
    if (theme === 'cyberpunk') return setTheme('dark');
    setTheme(theme === 'dark' ? 'light' : 'dark');
  }

  function toggleCyberpunk() {
    setTheme(currentTheme() === 'cyberpunk' ? 'dark' : 'cyberpunk', { toast: true });
  }

  function mountThemeButton() {
    if (document.getElementById('themeToggle')) return;
    const button = document.createElement('button');
    button.type = 'button';
    button.id = 'themeToggle';
    button.className = 'theme-toggle';
    button.addEventListener('click', toggleNormalTheme);
    document.body.appendChild(button);
    updateButton();
  }

  function installSecretTriggers() {
    // Hidden keyboard shortcut: Alt + Shift + K.
    document.addEventListener('keydown', event => {
      if (event.altKey && event.shiftKey && event.code === 'KeyK') {
        event.preventDefault();
        toggleCyberpunk();
      }
    });

    // Hidden landing-page easter egg: click the version number five times quickly.
    const version = document.querySelector('.version-mark');
    if (version) {
      let clicks = 0;
      let resetTimer = null;
      version.addEventListener('click', () => {
        clicks += 1;
        window.clearTimeout(resetTimer);
        resetTimer = window.setTimeout(() => { clicks = 0; }, 2200);
        if (clicks >= 5) {
          clicks = 0;
          window.clearTimeout(resetTimer);
          toggleCyberpunk();
        }
      });
    }
  }

  // The small inline boot script in each template applies the stored theme before
  // paint; this is a second safety pass for pages opened directly.
  const initial = storedTheme();
  if (initial !== 'light') root.dataset.theme = initial;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      mountThemeButton();
      installSecretTriggers();
    }, { once: true });
  } else {
    mountThemeButton();
    installSecretTriggers();
  }
})();
