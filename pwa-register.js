(() => {
  if (!('serviceWorker' in navigator)) return;

  const registrar = () => {
    navigator.serviceWorker
      .register('./sw.js', { scope: './' })
      .catch((erro) => {
        // A instalação PWA não deve interferir no funcionamento normal do sistema.
        console.info('Perfinorte: não foi possível registrar o modo instalável.', erro);
      });
  };

  if (document.readyState === 'complete') {
    registrar();
  } else {
    window.addEventListener('load', registrar, { once: true });
  }
})();