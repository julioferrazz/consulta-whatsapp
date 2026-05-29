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

# Garante que config.py seja encontrado mesmo rodando de outro diretorio
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import (
    REDIS_HOST, REDIS_PORT, REDIS_DB, REDIS_PASSWORD,
    NODE_API_URL, NUM_SESSIONS,
    MAX_RPS, BATCH_SIZE, BATCH_DELAY,
    QUEUE_KEY, RESULTS_KEY, COUNTER_KEY, RPS_PREFIX,
)


# ============================================================================
# Classe: GerenciadorRedis
# Responsabilidade: encapsula a conexao e operacoes atomicas no Redis
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
        """Remove todas as chaves de uma execucao anterior."""
        self._client.delete(QUEUE_KEY, RESULTS_KEY, COUNTER_KEY)

    def enfileirar(self, numero: str):
        self._client.rpush(QUEUE_KEY, numero)

    def desenfileirar(self) -> str | None:
        return self._client.lpop(QUEUE_KEY)

    def tamanho_fila(self) -> int:
        return self._client.llen(QUEUE_KEY)

    def reenfileirar(self, numero: str):
        """Coloca o numero de volta no final da fila (em caso de falha)."""
        self._client.rpush(QUEUE_KEY, numero)

    def adicionar_resultado(self, numero: str):
        self._client.sadd(RESULTS_KEY, numero)

    def obter_resultados(self) -> set:
        return self._client.smembers(RESULTS_KEY)

    def incrementar_contador(self) -> int:
        return int(self._client.incr(COUNTER_KEY))

    def contador_atual(self) -> int:
        return int(self._client.get(COUNTER_KEY) or 0)

    def adquirir_slot_rps(self):
        """
        Controle de taxa: bloqueia ate haver slot na janela de 1 segundo.
        Garante no maximo MAX_RPS requisicoes por segundo no total,
        usando uma chave Redis por segundo como contador atomico.
        """
        while True:
            agora = time.time()
            seg   = int(agora)
            chave = f'{RPS_PREFIX}:{seg}'

            pipe = self._client.pipeline()
            pipe.incr(chave)
            pipe.expire(chave, 2)   # expira apos 2s para nao acumular lixo
            count, _ = pipe.execute()

            if count <= MAX_RPS:
                return  # slot disponivel

            # Aguarda o restante do segundo atual antes de tentar novamente
            espera = 1.0 - (agora - seg) + 0.02
            time.sleep(max(0.05, espera))


# ============================================================================
# Classe: CarregadorExcel
# Responsabilidade: ler o Excel de entrada e normalizar os numeros
# ============================================================================
class CarregadorExcel:

    def __init__(self, caminho: str, gerenciador: GerenciadorRedis):
        self._caminho    = caminho
        self._gerenciador = gerenciador

    def carregar(self) -> int:
        """
        Le a planilha a partir da linha 2 (linha 1 = cabecalho),
        normaliza os numeros (apenas digitos) e os enfileira no Redis.
        Retorna a quantidade de numeros carregados.
        """
        self._gerenciador.limpar_estado()

        wb = openpyxl.load_workbook(self._caminho, read_only=True, data_only=True)
        ws = wb.active

        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            raw    = str(row[0]).strip()
            numero = ''.join(c for c in raw if c.isdigit())
            if numero:
                self._gerenciador.enfileirar(numero)
                count += 1

        wb.close()
        print(f'[Loader] {count} numeros carregados na fila Redis.')
        return count


# ============================================================================
# Classe: SalvadorExcel
# Responsabilidade: gravar o Excel de saida com os numeros validos
# ============================================================================
class SalvadorExcel:

    def __init__(self, caminho: str, gerenciador: GerenciadorRedis):
        self._caminho     = caminho
        self._gerenciador = gerenciador

    def salvar(self):
        resultados = sorted(self._gerenciador.obter_resultados())

        os.makedirs(os.path.dirname(os.path.abspath(self._caminho)), exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'WhatsApp Validos'
        ws['A1'] = 'Telefone'

        for i, numero in enumerate(resultados, start=2):
            ws[f'A{i}'] = numero

        wb.save(self._caminho)
        print(f'[Resultado] {len(resultados)} numeros com WhatsApp salvos em: {self._caminho}')


# ============================================================================
# Classe: Worker
# Responsabilidade: consumir a fila e consultar a API Node.js
# ============================================================================
class Worker:

    def __init__(self, session_id: int, gerenciador: GerenciadorRedis, stop_event: threading.Event):
        self._session_id  = session_id
        self._gerenciador = gerenciador
        self._stop_event  = stop_event

    def executar(self):
        print(f'[Worker {self._session_id}] Iniciado.')

        while not self._stop_event.is_set():
            numero = self._gerenciador.desenfileirar()

            # Fila vazia -> encerra
            if numero is None:
                if self._gerenciador.tamanho_fila() == 0:
                    break
                time.sleep(0.5)
                continue

            # ── Adquire slot de rate-limit ────────────────────────────────
            self._gerenciador.adquirir_slot_rps()

            # ── Chama a API Node.js ───────────────────────────────────────
            self._consultar(numero)

            delay = random.uniform(1.5, 4.0)
            time.sleep(delay)

            # ── Incrementa contador e verifica delay de lote ─────────────
            total = self._gerenciador.incrementar_contador()
            self._verificar_delay_lote(total)

        print(f'[Worker {self._session_id}] Encerrado.')

    def _consultar(self, numero: str):
        try:
            resp = requests.post(
                f'{NODE_API_URL}/check',
                json={'phone': numero, 'session': self._session_id},
                timeout=20,
            )

            if resp.status_code == 200:
                data = resp.json()
                if data.get('exists'):
                    self._gerenciador.adicionar_resultado(numero)
                    print(f'[Worker {self._session_id}]  {numero}  -> TEM WhatsApp')
                else:
                    print(f'[Worker {self._session_id}]  {numero}  -> sem WhatsApp')

            elif resp.status_code == 503:
                print(f'[Worker {self._session_id}] Sessao indisponivel, aguardando 3s...')
                self._gerenciador.reenfileirar(numero)
                time.sleep(3)

            else:
                print(f'[Worker {self._session_id}] HTTP {resp.status_code} para {numero} — recolocando na fila')
                self._gerenciador.reenfileirar(numero)

        except requests.exceptions.Timeout:
            print(f'[Worker {self._session_id}] Timeout para {numero} — recolocando na fila')
            self._gerenciador.reenfileirar(numero)
            time.sleep(2)

        except requests.exceptions.ConnectionError:
            print(f'[Worker {self._session_id}] Servidor Node.js inacessivel — aguardando 5s...')
            self._gerenciador.reenfileirar(numero)
            time.sleep(5)

        except Exception as exc:
            print(f'[Worker {self._session_id}] Erro inesperado ({exc}) — recolocando {numero}')
            self._gerenciador.reenfileirar(numero)
            time.sleep(2)

    @staticmethod
    def _verificar_delay_lote(total: int):
        if total > 0 and total % BATCH_SIZE == 0:
            print(f'\n[Sistema] {total} consultas realizadas -> pausa de {BATCH_DELAY}s...\n')
            time.sleep(BATCH_DELAY)


# ============================================================================
# Classe: Orquestrador
# Responsabilidade: coordenar todo o fluxo de execucao
# ============================================================================
class Orquestrador:

    def __init__(self, arquivo_entrada: str, arquivo_saida: str):
        self._entrada     = arquivo_entrada
        self._saida       = arquivo_saida
        self._gerenciador = GerenciadorRedis()

    def executar(self):
        # 1. Carrega numeros no Redis
        carregador = CarregadorExcel(self._entrada, self._gerenciador)
        total = carregador.carregar()
        if total == 0:
            print('[Erro] Nenhum numero encontrado na planilha.')
            sys.exit(1)

        # 2. Aguarda sessoes prontas
        sessoes_prontas = self._aguardar_sessoes(min_prontas=1)
        num_workers     = min(NUM_SESSIONS, sessoes_prontas)

        # 3. Inicia workers em threads separadas
        stop_event = threading.Event()
        threads    = []
        for i in range(num_workers):
            worker = Worker(i, self._gerenciador, stop_event)
            t = threading.Thread(target=worker.executar, daemon=True)
            threads.append(t)
            t.start()

        # 4. Aguarda todos concluirem
        for t in threads:
            t.join()
        stop_event.set()

        # 5. Salva resultados
        SalvadorExcel(self._saida, self._gerenciador).salvar()

        total_verificado = self._gerenciador.contador_atual()
        print(f'\n[Sistema] Concluido! Total verificado: {total_verificado} numeros.')

    @staticmethod
    def _aguardar_sessoes(min_prontas: int = 1) -> int:
        print('\n[Sistema] Verificando sessoes do WhatsApp...')
        tentativa = 0
        while True:
            try:
                resp  = requests.get(f'{NODE_API_URL}/status', timeout=5)
                st    = resp.json()
                prontas = sum(1 for v in st.values() if v is True)
                print(f'[Sistema] {prontas}/{NUM_SESSIONS} sessoes prontas...', end='\r')
                if prontas >= min_prontas:
                    print(f'\n[Sistema] {prontas} sessao(oes) pronta(s). Iniciando workers...\n')
                    return prontas
            except Exception:
                if tentativa == 0:
                    print('[Sistema] Aguardando servidor Node.js iniciar...')
                tentativa += 1
            time.sleep(2)


# ============================================================================
# Entry Point
# ============================================================================
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Uso: python main.py <entrada.xlsx> [saida.xlsx]')
        sys.exit(1)

    entrada = sys.argv[1]
    saida   = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        '..', 'output', 'resultado_whatsapp.xlsx'
    )

    if not os.path.exists(entrada):
        print(f'[Erro] Arquivo nao encontrado: {entrada}')
        sys.exit(1)

    Orquestrador(entrada, saida).executar()
