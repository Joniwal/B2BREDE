/* Controles gerais compartilhados pelo painel e pela página de análises. */

document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll(".btn-encerrar-app").forEach((button) => {
    button.addEventListener("click", () => encerrarAplicativo(button));
  });
});

async function encerrarAplicativo(button) {
  const confirmado = window.confirm(
    "Deseja encerrar o aplicativo REDEB2B? O navegador poderá ser fechado em seguida."
  );
  if (!confirmado) return;

  const botoes = [...document.querySelectorAll(".btn-encerrar-app")];
  botoes.forEach((item) => {
    item.disabled = true;
  });
  button.innerHTML = '<span class="spinner-border spinner-border-sm me-1" aria-hidden="true"></span>Encerrando...';

  try {
    const response = await fetch("/api/shutdown", {
      method: "POST",
      headers: { Accept: "application/json" },
      cache: "no-store",
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || payload.ok === false) {
      throw new Error(payload.error || "Não foi possível encerrar o aplicativo.");
    }

    document.body.innerHTML = `
      <main class="min-vh-100 d-flex align-items-center justify-content-center bg-light p-3">
        <section class="bg-white rounded-4 shadow-sm border p-5 text-center" style="max-width:520px;">
          <i class="bi bi-check-circle text-success" style="font-size:3rem;"></i>
          <h1 class="h4 mt-3">Aplicativo encerrado</h1>
          <p class="text-muted mb-0">O servidor REDEB2B foi finalizado. Você já pode fechar esta aba.</p>
        </section>
      </main>`;
  } catch (error) {
    botoes.forEach((item) => {
      item.disabled = false;
    });
    button.innerHTML = '<i class="bi bi-power"></i> Encerrar aplicativo';
    if (typeof showAlert === "function") {
      showAlert(error.message, "danger");
    } else {
      window.alert(error.message);
    }
  }
}
