from datetime import datetime
import re

def formatar_data_db(data_str):
    try:
        # Converter de DD/MM/YYYY para YYYY-MM-DD
        return datetime.strptime(data_str, '%d/%m/%Y').date()
    except ValueError:
        return None
    

def formatar_valor(valor):
    try:
        valor_str = str(valor).strip()
        valor_str = valor_str.replace('.', '').replace(',', '.')
        return float(valor_str)
    except ValueError:
        return None
    
def formatar_valor_protheus(valor):
    # Verifica se o valor é numérico e se contém vírgula
    if isinstance(valor, str):
        valor = valor.replace(',', '.')  # Troca a vírgula por ponto
        if not re.match(r'^\d+(\.\d{1,2})?$', valor):  # Verifica se é um número válido
            return None
        return float(valor)
    elif isinstance(valor, (int, float)):
        return valor
    return None