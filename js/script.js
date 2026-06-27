document.addEventListener('DOMContentLoaded', function () {
  const burger = document.getElementById('burger');
  const navMenu = document.querySelector('.nav-menu');
  if (burger && navMenu) {
    burger.addEventListener('click', function () {
      navMenu.classList.toggle('open');
      burger.setAttribute('aria-expanded', navMenu.classList.contains('open'));
    });
    navMenu.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', function () {
        navMenu.classList.remove('open');
        burger.setAttribute('aria-expanded', 'false');
      });
    });
  }
});
