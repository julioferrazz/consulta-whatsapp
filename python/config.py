# config.py — Configuracoes centrais do WhatsApp Checker

# ── Redis ────────────────────────────────────────────────────────────────────
REDIS_HOST = "localhost"
REDIS_PORT = 6380
REDIS_DB = 0
REDIS_PASSWORD = None  # None = sem senha; coloque a senha aqui se houver

# ── API Node.js ──────────────────────────────────────────────────────────────
NODE_API_URL = "http://localhost:3000"
NUM_SESSIONS = 2  # deve ser igual ao NUM_SESSIONS no index.js

# ── Controle de taxa ─────────────────────────────────────────────────────────
MAX_RPS = 1  # maximo de consultas por segundo (total, todos workers)
BATCH_SIZE = 30  # quantidade de consultas ate o delay obrigatorio
BATCH_DELAY = 35  # segundos de pausa a cada BATCH_SIZE consultas

# ── Chaves Redis ─────────────────────────────────────────────────────────────
QUEUE_KEY = "wa:queue"  # lista (FIFO) de numeros pendentes
RESULTS_KEY = "wa:results"  # set de numeros confirmados no WhatsApp
COUNTER_KEY = "wa:counter"  # contador global de consultas realizadas
RPS_PREFIX = "wa:rps"  # prefixo para janelas de rate-limit por segundo
MAX_RETRIES = (
    3  # numero maximo de tentativas para processar um numero antes de desistir
)
