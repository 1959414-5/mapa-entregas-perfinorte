(() => {
  const registrarServiceWorker = () => {
    if (!('serviceWorker' in navigator)) return;
    navigator.serviceWorker
      .register('./sw.js', { scope: './' })
      .catch((erro) => {
        // A instalação PWA não deve interferir no funcionamento normal do sistema.
        console.info('Perfinorte: não foi possível registrar o modo instalável.', erro);
      });
  };

  const configurarInstalacao = () => {
    const botao = document.getElementById('btnInstalarApp');
    if (!botao) return;

    const emStandalone = () =>
      window.matchMedia('(display-mode: standalone)').matches ||
      window.navigator.standalone === true;

    if (emStandalone()) {
      botao.hidden = true;
      return;
    }

    const suportaWebInstall =
      typeof Navigator !== 'undefined' &&
      typeof Navigator.prototype.install === 'function';

    let promptAdiado = null;

    // Chrome 158+: usa a interface nativa do navegador pela Web Install API.
    if (suportaWebInstall) botao.hidden = false;

    // Fallback para versões do Chrome que ainda usam beforeinstallprompt.
    window.addEventListener('beforeinstallprompt', (evento) => {
      evento.preventDefault();
      promptAdiado = evento;
      botao.hidden = false;
    });

    botao.addEventListener('click', async () => {
      if (botao.getAttribute('aria-disabled') === 'true') return;
      botao.setAttribute('aria-disabled', 'true');

      try {
        if (suportaWebInstall) {
          // A chamada acontece diretamente no clique para preservar a ativação do usuário.
          await navigator.install();
          botao.hidden = true;
          return;
        }

        if (promptAdiado) {
          const evento = promptAdiado;
          promptAdiado = null;
          await evento.prompt();
          const escolha = await evento.userChoice;

          if (escolha && escolha.outcome === 'accepted') {
            botao.hidden = true;
          } else if (!promptAdiado) {
            botao.hidden = true;
            const toast = document.getElementById('toast');
            if (toast) {
              toast.textContent = 'Instalação cancelada. Você pode tentar novamente pelo navegador.';
              toast.classList.add('show');
              window.setTimeout(() => toast.classList.remove('show'), 3500);
            }
          }
        }
      } catch (erro) {
        console.warn('Perfinorte: falha ao iniciar instalação.', erro);
        const toast = document.getElementById('toast');
        const mensagem = erro && erro.name === 'AbortError'
          ? 'Instalação cancelada. Você pode tentar novamente.'
          : 'O Chrome não conseguiu iniciar a instalação. Verifique se ela já está ativa ou permitida.';
        if (toast) {
          toast.textContent = mensagem;
          toast.classList.add('show');
          window.setTimeout(() => toast.classList.remove('show'), 4500);
        }
        if (!suportaWebInstall && !promptAdiado) botao.hidden = true;
      } finally {
        botao.removeAttribute('aria-disabled');
      }
    });

    window.addEventListener('appinstalled', () => {
      botao.hidden = true;
      const toast = document.getElementById('toast');
      if (toast) {
        toast.textContent = 'Perfinorte instalado com sucesso.';
        toast.classList.add('show');
        window.setTimeout(() => toast.classList.remove('show'), 3500);
      }
    });
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      registrarServiceWorker();
      configurarInstalacao();
    }, { once: true });
  } else {
    registrarServiceWorker();
    configurarInstalacao();
  }
})();