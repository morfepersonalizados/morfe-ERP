// Service Worker do Morfê ERP — cache mínimo, só pra deixar o sistema
// instalável e abrir rápido. NUNCA guarda dados da planilha, senha ou
// token — isso sempre vai direto pra rede, nunca fica salvo aqui.
const CACHE_NOME = 'morfe-erp-v1';
const ARQUIVOS_APP_SHELL = ['./', './index.html'];

self.addEventListener('install', (evento) => {
  self.skipWaiting();
  evento.waitUntil(
    caches.open(CACHE_NOME).then((cache) => cache.addAll(ARQUIVOS_APP_SHELL))
  );
});

self.addEventListener('activate', (evento) => {
  evento.waitUntil(
    caches.keys()
      .then((nomes) => Promise.all(nomes.filter((n) => n !== CACHE_NOME).map((n) => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (evento) => {
  const req = evento.request;

  // Só intercepta pedidos de abrir a página (navegação) e o próprio
  // index.html. Tudo mais — Google Sheets, fontes, scripts externos —
  // vai direto pra rede, sem passar pelo cache, sempre.
  const ehNavegacao = req.mode === 'navigate';
  const ehIndexHtml = req.url.endsWith('/index.html') || req.url.endsWith('/');

  if (ehNavegacao || ehIndexHtml) {
    evento.respondWith(
      // "Network first": tenta buscar a versão mais nova da internet
      // primeiro. Só usa o que está guardado se estiver sem internet —
      // assim nunca fica preso numa versão antiga do sistema à toa.
      fetch(req)
        .then((resposta) => {
          const copia = resposta.clone();
          caches.open(CACHE_NOME).then((cache) => cache.put(req, copia));
          return resposta;
        })
        .catch(() => caches.match(req).then((r) => r || caches.match('./index.html')))
    );
  }
  // Qualquer outro pedido (API do Google, fontes, etc) nem passa por aqui.
});
