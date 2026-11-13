"""Verifica somente autenticação e leitura da Caixa de entrada; não baixa anexos."""
from dotenv import load_dotenv
from automacao.config import BASE_DIR
from automacao.fonte_relatorios import FonteRelatorios, ErroEmail

def main():
    load_dotenv(BASE_DIR/'.env', override=False)
    fonte = None
    try:
        fonte = FonteRelatorios(BASE_DIR, BASE_DIR)
        if fonte.origem != 'gmail':
            raise ErroEmail('Defina AUTOMACAO_ORIGEM=gmail na .env para este teste.')
        fonte.conectar()
        print('Conexão com Gmail e acesso somente de leitura à Caixa de entrada: OK.')
        return 0
    except (ErroEmail, ValueError) as erro:
        print(f'Falha na verificação: {erro}')
        return 1
    finally:
        if fonte is not None:
            fonte.encerrar()

if __name__ == '__main__':
    raise SystemExit(main())
