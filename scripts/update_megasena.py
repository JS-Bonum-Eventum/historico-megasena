"""
update_megasena.py
Baixa o arquivo Excel da Mega-Sena do site da Caixa e converte para historico.txt
Formato de saída: uma linha por sorteio, 6 números separados por espaço, ordem crescente
"""

import requests
import openpyxl
from io import BytesIO

EXCEL_URL_CAIXA = "https://servicebus2.caixa.gov.br/portaldeloterias/api/resultados/download?modalidade=Mega-Sena"

OUTPUT_FILE = "historico.txt"
SEQUENCE_LENGTH = 6
MAX_NUMBER = 60

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    "Referer": "https://loterias.caixa.gov.br/Paginas/Mega-Sena.aspx",
}

def download_excel():
    print("Baixando Excel da Mega-Sena...")
    resp = requests.get(EXCEL_URL_CAIXA, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    print(f"Download OK: {len(resp.content)} bytes")
    return BytesIO(resp.content)

def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(file_bytes, read_only=True, data_only=True)
    ws = wb.active
    sequences = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        try:
            nums = []
            for i in range(1, 7):  # 6 números
                val = row[i]
                if val is not None:
                    n = int(val)
                    if 1 <= n <= MAX_NUMBER:
                        nums.append(n)
            if len(nums) == SEQUENCE_LENGTH:
                sorted_nums = sorted(set(nums))
                if len(sorted_nums) == SEQUENCE_LENGTH:
                    sequences.append(sorted_nums)
        except (ValueError, TypeError, IndexError):
            continue

    wb.close()
    print(f"Sequências válidas extraídas: {len(sequences)}")
    return sequences

def write_output(sequences):
    seen = set()
    unique = []
    for s in sequences:
        key = ','.join(map(str, s))
        if key not in seen:
            seen.add(key)
            unique.append(s)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        for seq in unique:
            f.write(' '.join(map(str, seq)) + '\n')

    print(f"historico.txt gerado: {len(unique)} linhas")

def main():
    file_bytes = download_excel()
    sequences = parse_excel(file_bytes)
    if not sequences:
        raise ValueError("Nenhuma sequência válida encontrada no Excel!")
    write_output(sequences)
    print("✅ Mega-Sena atualizado com sucesso!")

if __name__ == "__main__":
    main()
