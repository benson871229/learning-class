// Navbar shadow on scroll
window.addEventListener('scroll', () => {
  const nav = document.getElementById('mainNav');
  if (nav) nav.classList.toggle('shadow', window.scrollY > 10);
});
