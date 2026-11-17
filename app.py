import pandas as pd
import openpyxl

df = pd.read_excel("dados-educacao.xlsx")


def _ler_bruto(origem):
    """Lê CSV (utf-8 ou latin-1) ou Excel, sem cabeçalho."""
    nome = origem if isinstance(origem, str) else origem.name
    nome = nome.lower()

    if nome.endswith(".csv"):
        for enc in ("utf-8", "latin-1"):
            try:
                if not isinstance(origem, str):
                    origem.seek(0)
                return pd.read_csv(origem, header=None, encoding=enc, sep=None, engine="python")
            except UnicodeDecodeError:
                continue
        raise ValueError("Não foi possível ler o CSV (encoding).")

    if not isinstance(origem, str):
        origem.seek(0)
    if nome.endswith(".xls"):
        return pd.read_excel(origem, header=None, engine="xlrd")
    return pd.read_excel(origem, header=None)


homens = 0
mulheres = 0

for row in sheet.iter_rows(min_row=2, values_only=True):

  linha_texto = [str(celula).strip() for celula in row if celula is not None]


  if 'Mestrado' in linha_texto:

    if 'Masculino' in linha_texto or 'Homem' in linha_texto or 'M' in linha_texto:
      homens += 1
    elif (
        'Feminino' in linha_texto or 'Mulher' in linha_texto or 'F' in linha_texto
    ):
      mulheres += 1

print(f'Total com Mestrado - Homens: {homens}')
print(f'Total com Mestrado - Mulheres: {mulheres}')

