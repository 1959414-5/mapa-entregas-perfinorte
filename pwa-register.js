(() => {
  const registrarServiceWorker = () => {
    if (!('serviceWorker' in navigator)) return;
    navigator.serviceWorker
      .register('./sw.js', { scope: './' })
      .catch((erro) => {
        console.info('Perfinorte: não foi possível registrar o modo instalável.', erro);
      });
  };

  const configurarInstalacao = () => {
    const botao = document.getElementById('btnInstalarApp');
    if (!botao) return;

    const emStandalone = () =>
      window.matchMedia('(display-mode: standalone)').matches ||
      window.navigator.standalone === true;

    const supportsWebInstall =
      typeof Navigator !== 'undefined' &&
      typeof Navigator.prototype.install === 'function';

    let promptAdiado = null;

    if (emStandalone()) {
      botao.hidden = true;
      return;
    }

    if (supportsWebInstall) botao.hidden = false;

    window.addEventListener('beforeinstallprompt', (evento) => {
      evento.preventDefault();
      promptAdiado = evento;
      if (!emStandalone()) botao.hidden = false;
    });

    const mostrarMensagem = (mensagem) => {
      const toast = document.getElementById('toast');
      if (!toast) return;
      toast.textContent = mensagem;
      toast.classList.add('show');
      window.setTimeout(() => toast.classList.remove('show'), 5000);
    };

    const executarPromptAlternativo = async () => {
      if (!promptAdiado) return false;
      const evento = promptAdiado;
      promptAdiado = null;
      await evento.prompt();
      const escolha = await evento.userChoice;
      if (escolha && escolha.outcome === 'accepted') {
        botao.hidden = true;
        mostrarMensagem('Perfinorte instalado com sucesso.');
        return true;
      }
      botao.hidden = !promptAdiado;
      return false;
    };

    botao.addEventListener('click', async () => {
      if (botao.getAttribute('aria-disabled') === 'true') return;
      botao.setAttribute('aria-disabled', 'true');

      try {
        if (supportsWebInstall) {
          try {
            // Chamada direta no clique para preservar a ativação do usuário.
            await navigator.install();
            botao.hidden = true;
            mostrarMensagem('Perfinorte instalado ou já disponível como aplicativo.');
            return;
          } catch (erroWebInstall) {
            console.warn('Perfinorte: Web Install API falhou.', {
              name: erroWebInstall && erroWebInstall.name,
              message: erroWebInstall && erroWebInstall.message,
              error: erroWebInstall
            });
            if (promptAdiado) {
              await executarPromptAlternativo();
              return;
            }
            throw erroWebInstall;
          }
        }

        if (promptAdiado) {
          await executarPromptAlternativo();
          return;
        }

        botao.hidden = true;
        mostrarMensagem('O Chrome ainda não liberou a instalação. Recarregue o site e tente novamente pelo menu do navegador.');
      } catch (erro) {
        console.error('Perfinorte: erro ao instalar PWA.', erro);
        const nome = erro && erro.name ? erro.name : 'Erro desconhecido';
        const mensagens = {
          AbortError: 'O Chrome cancelou a instalação ou informou que o app já está instalado.',
          DataError: 'O Chrome rejeitou os dados do aplicativo. Verifique o manifesto e os ícones.',
          NotAllowedError: 'O Chrome bloqueou a instalação por falta de permissão ou ativação do usuário.',
          SecurityError: 'A política de segurança do navegador bloqueou a instalação.',
          TypeError: 'O Chrome não reconheceu os dados de instalação do aplicativo.',
          InvalidStateError: 'A instalação não pode ser iniciada nesta página.'
        };
        mostrarMensagem((mensagens[nome] || 'Não foi possível iniciar a instalação.') + ' (' + nome + ')');
      } finally {
        botao.removeAttribute('aria-disabled');
      }
    });

    window.addEventListener('appinstalled', () => {
      botao.hidden = true;
      mostrarMensagem('Perfinorte instalado com sucesso.');
    });
  };

  const inicializar = () => {
    registrarServiceWorker();
    configurarInstalacao();
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', inicializar, { once: true });
  } else {
    inicializar();
  }
})();