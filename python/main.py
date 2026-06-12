#!/usr/bin/env python3
"""
main.py — Orquestrador principal do WhatsApp Checker

Uso:
    python main.py <arquivo_entrada.xlsx> [arquivo_saida.xlsx]

Exemplo:
    python main.py ../input/numeros.xlsx ../output/resultado.xlsx

O script faz tudo em ordem:
    1. Le o Excel e carrega os numeros no Redis
    2. Aguarda o servidor Node.js ter ao menos 1 sessao pronta
    3. Inicia os workers (1 por sessao disponivel)
    4. Aguarda todos os workers terminarem
    5. Salva os numeros com WhatsApp em um novo Excel
"""

import sys
import os
import random
import time
import threading

import redis
import openpyxl
import requests

# ── Controle de pausa global ─────────────────────────────────────────────────
#
# RUNNING_EVENT:
#   set()   = sistema rodando normalmente  (estado inicial)
#   clear() = pausa ativa — todos os workers bloqueiam em .wait()
#
#
# _sleep_interruptivel():
#   Substitui time.sleep(delay) para que o delay seja abortado
#   assim que uma pausa começar, sem esperar ele terminar.
#
# ─────────────────────────────────────────────────────────────────────────────
# ── Controle de pausa global ─────────────────────────────────────────────────
RUNNING_EVENT = threading.Event()
RUNNING_EVENT.set()  # começa em estado "rodando"

# NOVAS VARIÁVEIS PARA SINCRONIZAÇÃO EXATA DE LOTE:
BATCH_LOCK = threading.Lock()
DISPATCH_COUNT = 0  # Conta quantas consultas iniciaram no lote atual
COMPLETED_COUNT = 0  # Conta quantas consultas terminaram no lote atual
# ─────────────────────────────────────────────────────────────────────────────


def _sleep_interruptivel(segundos: float) -> None:
    """
    Dorme por `segundos`, mas interrompe imediatamente se o RUNNING_EVENT
    for limpo (pausa iniciada), sem esperar o tempo restante.
    """
    prazo = time.time() + segundos
    while time.time() < prazo:
        if not RUNNING_EVENT.is_set():
            return  # pausa iniciada — sai do sleep antes do tempo
        time.sleep(0.05)  # granularidade de 50 ms


# Garante que config.py seja encontrado mesmo rodando de outro diretorio
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import (
    REDIS_HOST,
    REDIS_PORT,
    REDIS_DB,
    REDIS_PASSWORD,
    NODE_API_URL,
    NUM_SESSIONS,
    MAX_RPS,
    BATCH_SIZE,
    BATCH_DELAY,
    QUEUE_KEY,
    RESULTS_KEY,
    COUNTER_KEY,
    RPS_PREFIX,
    MAX_RETRIES,
)


# ============================================================================
# Classe: GerenciadorRedis
# ============================================================================
class GerenciadorRedis:

    def __init__(self):
        self._client = redis.Redis(
            host=REDIS_HOST,
            port=REDIS_PORT,
            db=REDIS_DB,
            password=REDIS_PASSWORD,
            decode_responses=True,
        )

    @property
    def client(self) -> redis.Redis:
        return self._client

    def limpar_estado(self):
        self._client.delete(QUEUE_KEY, RESULTS_KEY, COUNTER_KEY)

    def enfileirar(self, numero: str):
        self._client.rpush(QUEUE_KEY, numero)

    def tamanho_fila(self) -> int:
        return self._client.llen(QUEUE_KEY)

    def reenfileirar(self, numero: str, tentativas: int):
        if tentativas < MAX_RETRIES:
            nova = tentativas + 1
            self._client.rpush(QUEUE_KEY, f"{numero}:{nova}")
            print(f"[Redis] {numero} reenfileirado (tentativa {nova}/{MAX_RETRIES})")
        else:
            print(f"[Redis] {numero} descartado após {MAX_RETRIES} tentativas")

    def adicionar_resultado(self, numero: str):
        self._client.sadd(RESULTS_KEY, numero)

    def obter_resultados(self) -> set:
        return self._client.smembers(RESULTS_KEY)

    def contador_atual(self) -> int:
        return int(self._client.get(COUNTER_KEY) or 0)

    def adquirir_slot_rps(self):
        """Bloqueia até haver slot disponível na janela de 1 segundo."""
        while True:
            agora = time.time()
            seg = int(agora)
            chave = f"{RPS_PREFIX}:{seg}"

            pipe = self._client.pipeline()
            pipe.incr(chave)
            pipe.expire(chave, 2)
            count, _ = pipe.execute()

            if count <= MAX_RPS:
                return

            espera = 1.0 - (agora - seg) + 0.02
            time.sleep(max(0.05, espera))


# ============================================================================
# Classe: CarregadorExcel
# ============================================================================
class CarregadorExcel:

    def __init__(self, caminho: str, gerenciador: GerenciadorRedis):
        self._caminho = caminho
        self._gerenciador = gerenciador

    def carregar(self) -> int:
        self._gerenciador.limpar_estado()

        wb = openpyxl.load_workbook(self._caminho, read_only=True, data_only=True)
        ws = wb.active

        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            raw = str(row[0]).strip()
            numero = "".join(c for c in raw if c.isdigit())
            if numero:
                self._gerenciador.enfileirar(f"55{numero}:0")
                count += 1

        wb.close()
        print(f"[Loader] {count} numeros carregados na fila Redis.")
        return count


# ============================================================================
# Classe: SalvadorExcel
# ============================================================================
class SalvadorExcel:

    def __init__(self, caminho: str, gerenciador: GerenciadorRedis):
        self._caminho = caminho
        self._gerenciador = gerenciador

    def salvar(self):
        resultados = sorted(self._gerenciador.obter_resultados())

        os.makedirs(os.path.dirname(os.path.abspath(self._caminho)), exist_ok=True)

        # Abre arquivo existente ou cria um novo
        if os.path.exists(self._caminho):
            wb = openpyxl.load_workbook(self._caminho)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "WhatsApp Validos"
            ws["A1"] = "Telefone"

        # Primeira linha vazia da coluna A
        linha = 2
        while ws[f"A{linha}"].value not in (None, ""):
            linha += 1

        # Salva novos números sem o DDI 55
        adicionados = 0

        for numero in resultados:
            numero_sem_ddi = numero[2:] if numero.startswith("55") else numero

            ws[f"A{linha}"] = numero_sem_ddi
            linha += 1
            adicionados += 1

        wb.save(self._caminho)

        print(f"[Resultado] {adicionados} numeros adicionados em: {self._caminho}")


# ============================================================================
# Classe: Worker
# ============================================================================
class Worker:

    def __init__(
        self,
        session_id: int,
        gerenciador: GerenciadorRedis,
        stop_event: threading.Event,
        total: int,
    ):
        self._session_id = session_id
        self._gerenciador = gerenciador
        self._stop_event = stop_event
        self._total = total

    def executar(self):
        global DISPATCH_COUNT, COMPLETED_COUNT
        print(f"[Worker {self._session_id}] Iniciado.")

        while not self._stop_event.is_set():

            # 1. Aguarda pausa geral, se ativa
            RUNNING_EVENT.wait()

            # ── BARREIRA DE LOTE EXATO ──────────────────────────────────────────
            with BATCH_LOCK:
                if DISPATCH_COUNT >= BATCH_SIZE:
                    # O limite foi atingido. Aguarda os workers em andamento
                    # terminarem suas consultas deste lote para não ter "sobras".
                    while COMPLETED_COUNT < BATCH_SIZE:
                        if self._stop_event.is_set():
                            break
                        # Libera o lock rápido para outros workers conseguirem registrar a conclusão
                        BATCH_LOCK.release()
                        time.sleep(0.1)
                        BATCH_LOCK.acquire()

                    if self._stop_event.is_set():
                        break

                    # Trava os workers no início do loop
                    RUNNING_EVENT.clear()

                    total_atual = self._gerenciador.contador_atual()
                    print(
                        f"\n[Sistema] {total_atual} consultas realizadas -> pausa de {BATCH_DELAY}s...\n"
                    )

                    time.sleep(
                        BATCH_DELAY
                    )  # Faz a pausa exata sem ninguém consultar nada

                    # Reseta os contadores para o próximo ciclo
                    DISPATCH_COUNT = 0
                    COMPLETED_COUNT = 0
                    RUNNING_EVENT.set()  # Libera a continuação
                    print("\n[Sistema] Pausa encerrada. Retomando...\n")

                # Registra que este worker vai despachar uma nova consulta agora
                DISPATCH_COUNT += 1
            # ────────────────────────────────────────────────────────────────────

            # ── 2. Pega próximo número da fila ──────────────────────────────────
            item = self._gerenciador.client.blpop(QUEUE_KEY, timeout=2)
            if item is None:
                # Fila vazia → reverte o despacho e encerra worker
                with BATCH_LOCK:
                    DISPATCH_COUNT -= 1
                break

            numero_com_tentativas = item[1]
            numero, tentativas = numero_com_tentativas.split(":")
            tentativas = int(tentativas)

            # ── 3. Rate-limit global ────────────────────────────────────────────
            self._gerenciador.adquirir_slot_rps()

            # ── 4. Consulta API Node.js ─────────────────────────────────────────
            self._consultar(numero, tentativas)

            # ── 5. Incrementa contador global (atômico no Redis) ────────────────
            total = self._gerenciador.client.incr(COUNTER_KEY)
            restantes = max(0, self._total - total)
            print(
                f"[Worker {self._session_id}] Progresso: {total}/{self._total}"
                f"  | {restantes} Restantes"
            )

            # ── MARCA CONCLUSÃO ─────────────────────────────────────────────────
            # Avisa a barreira do lote que essa consulta finalizou 100%
            with BATCH_LOCK:
                COMPLETED_COUNT += 1

            # ── 6. Delay interruptível ──────────────────────────────────────────
            _sleep_interruptivel(random.uniform(5.0, 7.5))

        print(f"[Worker {self._session_id}] Encerrado.")

    def _consultar(self, numero: str, tentativas: int):
        try:
            resp = requests.post(
                f"{NODE_API_URL}/check",
                json={"phone": numero, "session": self._session_id},
                timeout=20,
            )

            if resp.status_code == 200:
                data = resp.json()
                if data.get("exists"):
                    self._gerenciador.adicionar_resultado(numero)
                    print(f"[Worker {self._session_id}]  {numero}  -> TEM WhatsApp")
                else:
                    print(f"[Worker {self._session_id}]  {numero}  -> sem WhatsApp")

            elif resp.status_code == 503:
                print(
                    f"[Worker {self._session_id}] Sessao indisponivel, aguardando 3s..."
                )
                self._gerenciador.reenfileirar(numero, tentativas)
                time.sleep(3)

            else:
                print(
                    f"[Worker {self._session_id}] HTTP {resp.status_code}"
                    f" para {numero} — recolocando na fila"
                )
                self._gerenciador.reenfileirar(numero, tentativas)

        except requests.exceptions.Timeout:
            print(
                f"[Worker {self._session_id}] Timeout para {numero} — recolocando na fila"
            )
            self._gerenciador.reenfileirar(numero, tentativas)
            time.sleep(2)

        except requests.exceptions.ConnectionError:
            print(
                f"[Worker {self._session_id}] Servidor Node.js inacessivel — aguardando 5s..."
            )
            self._gerenciador.reenfileirar(numero, tentativas)
            time.sleep(5)

        except Exception as exc:
            print(
                f"[Worker {self._session_id}] Erro inesperado ({exc})"
                f" — recolocando {numero}"
            )
            self._gerenciador.reenfileirar(numero, tentativas)
            time.sleep(2)


# ============================================================================
# Classe: Orquestrador
# ============================================================================
class Orquestrador:

    def __init__(self, arquivo_entrada: str, arquivo_saida: str):
        self._entrada = arquivo_entrada
        self._saida = arquivo_saida
        self._gerenciador = GerenciadorRedis()

    def executar(self):
        # 1. Carrega números no Redis
        carregador = CarregadorExcel(self._entrada, self._gerenciador)
        total = carregador.carregar()
        if total == 0:
            print("[Erro] Nenhum numero encontrado na planilha.")
            sys.exit(1)

        # 2. Aguarda sessões prontas
        sessoes_prontas = self._aguardar_sessoes(min_prontas=1)
        num_workers = min(NUM_SESSIONS, sessoes_prontas)

        # 3. Inicia workers
        # stop_event só é usado para parada forçada (Ctrl+C).
        # Em operação normal cada worker encerra quando a fila esvazia.
        stop_event = threading.Event()
        threads = []
        for i in range(num_workers):
            worker = Worker(i, self._gerenciador, stop_event, total)
            t = threading.Thread(target=worker.executar, daemon=True)
            threads.append(t)
            t.start()

        # 4. Aguarda todos concluírem
        try:
            for t in threads:
                t.join()
        except KeyboardInterrupt:
            print("\n[Sistema] Interrompido pelo usuário. Encerrando...")
            stop_event.set()
            RUNNING_EVENT.set()  # desbloqueia workers presos em wait()
            for t in threads:
                t.join(timeout=10)

        # 5. Salva resultados
        SalvadorExcel(self._saida, self._gerenciador).salvar()

        total_verificado = self._gerenciador.contador_atual()
        print(f"\n[Sistema] Concluido! Total verificado: {total_verificado} numeros.")

    @staticmethod
    def _aguardar_sessoes(min_prontas: int = 1) -> int:
        print("\n[Sistema] Verificando sessoes do WhatsApp...")
        tentativa = 0
        while True:
            try:
                resp = requests.get(f"{NODE_API_URL}/status", timeout=5)
                st = resp.json()
                prontas = sum(1 for v in st.values() if v is True)
                print(
                    f"[Sistema] {prontas}/{NUM_SESSIONS} sessoes prontas...", end="\r"
                )
                if prontas >= min_prontas:
                    print(
                        f"\n[Sistema] {prontas} sessao(oes) pronta(s). Iniciando workers...\n"
                    )
                    return prontas
            except Exception:
                if tentativa == 0:
                    print("[Sistema] Aguardando servidor Node.js iniciar...")
                tentativa += 1
            time.sleep(2)


# ============================================================================
# Entry Point
# ============================================================================
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python main.py <entrada.xlsx> [saida.xlsx]")
        sys.exit(1)

    entrada = sys.argv[1]
    saida = (
        sys.argv[2]
        if len(sys.argv) > 2
        else os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "..",
            "output",
            "resultado_whatsapp.xlsx",
        )
    )

    if not os.path.exists(entrada):
        print(f"[Erro] Arquivo nao encontrado: {entrada}")
        sys.exit(1)

    Orquestrador(entrada, saida).executar()
