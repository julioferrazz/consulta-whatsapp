'use strict';

/**
 * index.js — Servidor WhatsApp Checker
 *
 * Gerencia sessoes do WhatsApp via Baileys e expoe uma API HTTP
 * para que o Python consulte se um numero possui conta no WhatsApp.
 *
 * Endpoints:
 *   GET  /status        -> estado de cada sessao { session_0: bool, ... }
 *   POST /check         -> { phone: "5511999999999", session: 0 }
 *                          retorna { phone, exists: bool }
 */

const {
  default: makeWASocket,
  DisconnectReason,
  useMultiFileAuthState,
  fetchLatestBaileysVersion,
} = require('@whiskeysockets/baileys');

const express = require('express');
const qrcode = require('qrcode-terminal');
const path = require('path');
const fs = require('fs');
const pino = require('pino');

// ---------------------------------------------------------------------------
// Configuracoes
// ---------------------------------------------------------------------------
const PORT = 3000;
const NUM_SESSIONS = 2;
const SESSIONS_DIR = path.join(__dirname, 'sessions');

// ---------------------------------------------------------------------------
// Estado global das sessoes
// { 0: { sock, ready: bool }, 1: {...}, 2: {...} }
// ---------------------------------------------------------------------------
const sessions = {};

fs.mkdirSync(SESSIONS_DIR, { recursive: true });

// Logger silencioso para o Baileys (evita spam no terminal)
const logger = pino({ level: 'silent' });

// ---------------------------------------------------------------------------
// Cria (ou reconecta) uma sessao pelo ID
// ---------------------------------------------------------------------------
async function startSession(id) {
  const authDir = path.join(SESSIONS_DIR, `session_${id}`);
  fs.mkdirSync(authDir, { recursive: true });

  const { state, saveCreds } = await useMultiFileAuthState(authDir);
  const { version } = await fetchLatestBaileysVersion();

  const sock = makeWASocket({
    version,
    auth: state,
    logger,
    browser: ['WA-Checker', 'Chrome', '120.0'],
    syncFullHistory: false,
  });

  sock.ev.on('creds.update', saveCreds);

  // Promise que será resolvida quando conectar
  let resolveConnected;
  const connectedPromise = new Promise((resolve) => {
    resolveConnected = resolve;
  });

  sock.ev.on('connection.update', ({ connection, lastDisconnect, qr }) => {

    // Quando QR aparecer
    if (qr) {
      console.clear();

      console.log('\n' + '='.repeat(55));
      console.log(`  SESSAO ${id} — Escaneie o QR Code abaixo`);
      console.log('='.repeat(55) + '\n');

      qrcode.generate(qr, { small: true });

      console.log('\n[Aguardando conexão por até 40 segundos...]');
    }

    if (connection === 'close') {
      const code =
        lastDisconnect?.error?.output?.statusCode ||
        lastDisconnect?.error?.output?.payload?.statusCode;

      const shouldReconnect = code !== DisconnectReason.loggedOut;

      sessions[id] = { sock: null, ready: false };

      if (shouldReconnect) {
        console.log(`[Sessao ${id}] Conexao encerrada (${code}). Reconectando em 5s...`);
        setTimeout(() => startSession(id), 5000);
      } else {
        console.log(`[Sessao ${id}] Deslogado. Delete a pasta sessions/session_${id} e reinicie.`);
      }

    } else if (connection === 'open') {
      console.log(`[Sessao ${id}] Conectado com sucesso!`);

      sessions[id].ready = true;

      // Libera imediatamente a próxima sessão
      resolveConnected();
    }
  });

  sessions[id] = { sock, ready: false };

  // Espera:
  // - conectar
  // OU
  // - 40 segundos
  await Promise.race([
    connectedPromise,
    new Promise(resolve => setTimeout(resolve, 40000))
  ]);
}

// ---------------------------------------------------------------------------
// Inicia todas as sessoes (com intervalo para nao sobrecarregar)
// ---------------------------------------------------------------------------
(async () => {
  console.log('\n[Servidor] Iniciando ' + NUM_SESSIONS + ' sessoes do WhatsApp...\n');

  for (let i = 0; i < NUM_SESSIONS; i++) {
    console.log(`\n[Servidor] Abrindo sessão ${i}...\n`);
    await startSession(i);
  }

  console.log('\n[Servidor] Todas as sessoes foram iniciadas.');
  console.log('[Servidor] API disponivel em http://localhost:' + PORT + '\n');
})();

// ---------------------------------------------------------------------------
// API Express
// ---------------------------------------------------------------------------
const app = express();
app.use(express.json());

/**
 * GET /status
 * Retorna o estado de cada sessao.
 * Exemplo de resposta: { "session_0": true, "session_1": false, "session_2": true }
 */
app.get('/status', (_req, res) => {
  const status = {};
  for (let i = 0; i < NUM_SESSIONS; i++) {
    status['session_' + i] = sessions[i]?.ready === true;
  }
  res.json(status);
});

/**
 * POST /check
 * Body:    { "phone": "5511999999999", "session": 0 }
 * Retorna: { "phone": "5511999999999", "exists": true }
 *
 * O campo "phone" deve conter apenas digitos, incluindo o codigo do pais.
 * Exemplo Brasil: 5511987654321
 */
app.post('/check', async (req, res) => {
  const { phone, session } = req.body;

  if (phone === undefined || session === undefined) {
    return res.status(400).json({ error: 'Campos "phone" e "session" sao obrigatorios.' });
  }

  const s = sessions[session];

  if (!s || !s.ready || !s.sock) {
    return res.status(503).json({ error: 'Sessao ' + session + ' nao esta pronta.' });
  }

  try {
    const jid = phone + '@s.whatsapp.net';
    const results = await s.sock.onWhatsApp(jid);
    const exists = Array.isArray(results) && results.length > 0 && results[0].exists === true;
    res.json({ phone, exists });
  } catch (err) {
    console.error('[Sessao ' + session + '] Erro ao verificar ' + phone + ': ' + err.message);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log('[API] Servidor HTTP escutando na porta ' + PORT);
});
