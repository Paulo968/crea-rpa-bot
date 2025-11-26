from automation.bot import executar_lote
import json

def processar_contratos(
    df,
    inicio,
    log_callback,
    numero_inicial=None,
    cpf_cnpj_inicial=None,
    data_registro_inicial=None,
    callback_atualizar_contrato=None  # <- NOVO, opcional
):
    try:
        # Passa o callback pra frente, se tiver
        executar_lote(
            df,
            inicio,
            log_callback,
            numero_inicial,
            cpf_cnpj_inicial,
            data_registro_inicial,
            callback_atualizar_contrato  # <- NOVO, pode ser None
        )
    except Exception as e:
        log_callback(f"❌ Erro geral no processamento: {e}")
