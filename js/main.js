// Portfolio BTS SIO SISR – Kylian Grafeille Clément

// Remplace une image absente par un emplacement "à ajouter" indiquant le chemin attendu
function toPlaceholder(img) {
  const ph = document.createElement('div');
  ph.className = 'ph ' + img.className;
  ph.innerHTML = '<strong>' + (img.dataset.ph || 'Image') + '</strong>' +
    '<span>Image à ajouter</span>' +
    '<code>' + img.getAttribute('src') + '</code>';
  img.replaceWith(ph);
}

document.querySelectorAll('img[data-ph]').forEach((img) => {
  if (img.complete && img.naturalWidth === 0) toPlaceholder(img);
  else img.addEventListener('error', () => toPlaceholder(img), { once: true });
});

// Menu mobile
const burger = document.querySelector('.burger');
const nav = document.querySelector('.nav');
burger.addEventListener('click', () => {
  const open = nav.classList.toggle('is-open');
  burger.setAttribute('aria-expanded', open);
});
nav.querySelectorAll('a').forEach((a) => a.addEventListener('click', () => {
  nav.classList.remove('is-open');
  burger.setAttribute('aria-expanded', 'false');
}));

// Lien actif dans la navigation selon la section visible
const navLinks = [...nav.querySelectorAll('a')];
const spy = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (!entry.isIntersecting) return;
    let id = entry.target.id;
    if (id === 'synthese') id = 'competences';
    navLinks.forEach((a) => a.classList.toggle('is-active', a.getAttribute('href') === '#' + id));
  });
}, { rootMargin: '-45% 0px -50% 0px' });
document.querySelectorAll('main section[id]').forEach((s) => spy.observe(s));

// Apparition progressive des blocs
const reveal = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      entry.target.classList.add('is-visible');
      reveal.unobserve(entry.target);
    }
  });
}, { threshold: 0.1 });
document.querySelectorAll('.section .container > *').forEach((el) => {
  el.classList.add('reveal');
  reveal.observe(el);
});

// Agrandissement des captures au clic
const lightbox = document.createElement('div');
lightbox.className = 'lightbox';
lightbox.innerHTML = '<img alt="">';
document.body.appendChild(lightbox);
document.addEventListener('click', (e) => {
  const img = e.target.closest('.veille__shot');
  if (img && img.tagName === 'IMG') {
    lightbox.querySelector('img').src = img.src;
    lightbox.classList.add('is-open');
  }
});
lightbox.addEventListener('click', () => lightbox.classList.remove('is-open'));
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') lightbox.classList.remove('is-open'); });

document.getElementById('year').textContent = new Date().getFullYear();
