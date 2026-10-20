(() => {
    const key = 'paper_theme';
    const root = document.documentElement;
    const buttons = document.querySelectorAll('[data-theme-toggle]');
    const apply = theme => {
        root.dataset.theme = theme;
        buttons.forEach(button => {
            button.textContent = theme === 'dark' ? '☀️ 日间模式' : '🌙 夜间模式';
            button.setAttribute('aria-label', theme === 'dark' ? '切换日间模式' : '切换夜间模式');
            button.setAttribute('aria-pressed', theme === 'dark' ? 'true' : 'false');
        });
    };
    apply(root.dataset.theme === 'dark' ? 'dark' : 'light');
    buttons.forEach(button => button.addEventListener('click', () => {
        const next = root.dataset.theme === 'dark' ? 'light' : 'dark';
        apply(next);
        try { localStorage.setItem(key, next); } catch (_) { /* Theme still works for this page. */ }
    }));
})();
