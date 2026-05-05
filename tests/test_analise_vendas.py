import pandas as pd
import pytest

from src.analise_vendas import (
    COLUNAS_OBRIGATORIAS,
    ErroArquivoEntrada,
    calcular_metricas,
    formatar_moeda,
    validar_colunas,
)


def test_formatar_moeda_usa_padrao_brasileiro():
    assert formatar_moeda(1234.5) == "R$ 1.234,50"


def test_validar_colunas_rejeita_planilha_incompleta():
    df = pd.DataFrame({"produto": ["Notebook"]})

    with pytest.raises(ErroArquivoEntrada):
        validar_colunas(df)


def test_calcular_metricas_com_base_valida():
    df = pd.DataFrame(
        [
            {
                "data_venda": pd.Timestamp("2026-01-10"),
                "produto": "Notebook",
                "categoria": "Eletronicos",
                "quantidade": 2,
                "preco_unitario": 3000.0,
                "cliente": "Cliente A",
                "cidade": "Manaus",
                "estado": "AM",
                "forma_pagamento": "Pix",
                "faturamento": 6000.0,
                "mes": "2026-01",
                "ano": 2026,
            },
            {
                "data_venda": pd.Timestamp("2026-01-11"),
                "produto": "Mouse",
                "categoria": "Acessorios",
                "quantidade": 3,
                "preco_unitario": 100.0,
                "cliente": "Cliente B",
                "cidade": "Manaus",
                "estado": "AM",
                "forma_pagamento": "Cartao",
                "faturamento": 300.0,
                "mes": "2026-01",
                "ano": 2026,
            },
        ]
    )

    assert COLUNAS_OBRIGATORIAS.issubset(df.columns)

    metricas = calcular_metricas(df)
    indicadores = metricas["indicadores"].set_index("indicador")["valor"]

    assert indicadores["Faturamento total"] == 6300.0
    assert indicadores["Quantidade total vendida"] == 5
    assert indicadores["Produto mais vendido"] == "Mouse"
