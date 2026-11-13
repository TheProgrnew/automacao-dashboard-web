"""Trava exclusiva de processo único, liberada mesmo se o processo terminar.

Funciona em Windows (msvcrt) e Unix (fcntl). O arquivo é apenas um marcador,
não contém dados e não deve ser excluído enquanto houver execução.
"""
from contextlib import contextmanager
from pathlib import Path
import os
from django.conf import settings

class ExecucaoOcupada(Exception):
    pass

@contextmanager
def bloquear_execucao():
    caminho = Path(settings.BASE_DIR) / "automacao_execucao.lock"
    with open(caminho, "a+b") as arquivo:
        arquivo.seek(0)
        if os.name == "nt":
            import msvcrt
            try:
                msvcrt.locking(arquivo.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc:
                raise ExecucaoOcupada("Já existe uma atualização em andamento.") from exc
            try:
                yield
            finally:
                arquivo.seek(0)
                msvcrt.locking(arquivo.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            try:
                fcntl.flock(arquivo.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError as exc:
                raise ExecucaoOcupada("Já existe uma atualização em andamento.") from exc
            try:
                yield
            finally:
                fcntl.flock(arquivo.fileno(), fcntl.LOCK_UN)
