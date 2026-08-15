// Service worker mínimo — permite que o navegador ofereça
// "Adicionar à tela inicial" para a página do credenciamento.
// Não faz cache agressivo, então sempre pega dados atualizados do servidor.

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  // Passa direto para a rede (sem cache) — o app depende de dados em tempo real.
  event.respondWith(fetch(event.request));
});
