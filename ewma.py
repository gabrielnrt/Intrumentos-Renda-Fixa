"""
A ideia desse programa eh usar o pandas pra calcular o ewma,
tanto na versao 'normal' quanto na do dacorogna.
"""

import numpy as np
from scipy import signal
import pandas as pd
import xlwings as xw


WORKBOOK = "teste-ewma.xlsm"
SHEET_NORMAL = "ewma-normal"
SHEET_DACOROGNA = "ewma-dacorogna"


if __name__ == "__main__":

    wb = xw.Book(WORKBOOK)

    # --------------------------------------------------------------------------------
    # DACOROGNA

    sheet_dacorogna = wb.sheets[SHEET_DACOROGNA]

    df_dacorogna = pd.DataFrame(sheet_dacorogna["c5"].expand("down").value)

    # nessa celula ta o valor do tau descrito no dacorogna
    tau = sheet_dacorogna["K4"].value

    # esse alpha eh o que aparece na documentacao do pandas, e a formula eh essa pra
    # poder conectar com a notacao do dacorogna
    alpha_pandas = 1 - np.exp(-1 / tau)

    # O ewma aproximado oferecido pelo python fazendo adjust=True
    df_dacorogna["ewma aprox dacorogna python"] = (
        df_dacorogna[0].ewm(alpha=alpha_pandas, adjust=True).mean()
    )

    # ja o ewma do dacorogna eh obtido fazendo adjust=False
    df_dacorogna["ewma dacorogna python"] = (
        df_dacorogna[0].ewm(alpha=alpha_pandas, adjust=False).mean()
    )

    # --------------------------------------------------------------------------------
    # EWMA NORMAL

    sheet_normal = wb.sheets[SHEET_NORMAL]

    # esse eh o tamanho da janela
    M = int(sheet_normal["I2"].value)

    # esse eh um parametro que aparece na documentacao do scipy, que serviria como 'centro' da janela,
    # supondo que a janela seria simetrica. Na doc esta melhor explicado.
    # ele precisa estar dessa forma pq se for c=0, os dados mais antigos teriam mais peso e os mais recentes menos peso.
    c = M - 1

    # esse eh o lambda, parametro entre 0 e 1, gerador de pesos
    lamb = sheet_normal["I3"].value

    # esse eh o tau que aparece na documentacao do scipy, nao do dacorogna,
    # embora possivelmente tenha uma documentacao
    tau_scipy = -1 / np.log(lamb)

    # esse eh o vetor de pesos gerados dados esses inputs
    w_vector = signal.windows.exponential(M=M, center=M - 1, tau=tau_scipy, sym=False)

    df_normal = pd.DataFrame(sheet_normal["c4"].expand("down").value)

    # ESSE EH O EWMA CALCULADO DE FORMA RAPIDA USANDO O SCIPY E PANDAS.
    df_normal["EWMA-PYTHON"] = (
        df_normal[0]
        .rolling(window=M, win_type="exponential")
        .mean(center=c, tau=tau_scipy, sym=False)
    )

    ...
