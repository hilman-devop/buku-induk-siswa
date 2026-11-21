/* Tema terang/gelap: dipilih sekali, diingat di localStorage.
   Dimuat di <head> supaya tidak ada kedip warna saat halaman dibuka. */
(function () {
  var KEY = 'tema', root = document.documentElement, saved = null;
  try { saved = localStorage.getItem(KEY); } catch (e) {}
  var fromUrl = new URLSearchParams(location.search).get('tema');
  if (fromUrl === 'gelap' || fromUrl === 'terang') { saved = fromUrl; }
  if (saved !== 'gelap' && saved !== 'terang') {
    saved = window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'gelap' : 'terang';
  }
  root.setAttribute('data-theme', saved);

  function bind() {
    var btn = document.getElementById('tema');
    if (!btn) return;
    btn.setAttribute('aria-pressed', String(root.getAttribute('data-theme') === 'gelap'));
    btn.onclick = function () {
      var next = root.getAttribute('data-theme') === 'gelap' ? 'terang' : 'gelap';
      root.setAttribute('data-theme', next);
      btn.setAttribute('aria-pressed', String(next === 'gelap'));
      btn.title = next === 'gelap' ? 'Pakai tema terang' : 'Pakai tema gelap';
      try { localStorage.setItem(KEY, next); } catch (e) {}
    };
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bind);
  } else { bind(); }
})();
