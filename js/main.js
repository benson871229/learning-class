document.addEventListener('DOMContentLoaded', () => {
  const btn  = document.getElementById('mobile-menu-btn');
  const menu = document.getElementById('mobile-menu');
  const openIcon  = document.getElementById('menu-open-icon');
  const closeIcon = document.getElementById('menu-close-icon');

  if (btn && menu) {
    btn.addEventListener('click', () => {
      const isHidden = menu.classList.toggle('hidden');
      openIcon?.classList.toggle('hidden', !isHidden);
      closeIcon?.classList.toggle('hidden', isHidden);
    });
  }

  // Highlight active nav link
  const links = document.querySelectorAll('nav a[href]');
  const current = location.pathname.split('/').pop() || 'index.html';
  links.forEach(a => {
    if (a.getAttribute('href') === current) a.classList.add('text-indigo-600', 'font-semibold');
  });
});
