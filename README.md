
# autofillPlant

Automação em Python + Selenium para atualizar o campo "Facility type" para "Plant Location" em registros do ServiceNow, iterando hostnames a partir da planilha Inventario_RAD.xlsx (aba "INVENTARIO RAD", coluna "HOSTNAME").

## Pré-requisitos
- Python 3.10+
- Pacotes: selenium, pyautogui, python-dotenv, pandas, openpyxl
- Chrome portátil na pasta chrome-win64/chrome.exe

## Instalação
```
python -m venv .venv
. .venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install selenium pyautogui python-dotenv pandas openpyxl
```

## Configuração
Crie um arquivo .env na raiz do projeto (existe um .env.sample para copiar):
```
INSTANCE_URL=https://SUA_INSTANCIA.service-now.com
SN_USER=seu.usuario
SN_PASS=sua.senha
EXCEL_PATH=Inventario_RAD.xlsx
EXCEL_SHEET=INVENTARIO RAD
EXCEL_COLUMN=HOSTNAME
FACILITY_TYPE=Plant Location
USE_COORDINATE_SAVE=false
RIGHT_CLICK_X=1328
RIGHT_CLICK_Y=190
CHROME_BINARY=chrome-win64/chrome.exe
# MAX_ROWS=5
```

## Execução
```
python .\sn_bulk_update_facility.py
```

Ao final, um relatório resultado_facility.csv será gerado com o status de cada hostname. Em caso de erros, screenshots error_<hostname>.png serão salvos.

## Observações
- Preferir salvar via DOM (USE_COORDINATE_SAVE=false) para maior robustez.
- Se usar coordenadas, mantenha o Windows com escala 100% e o Chrome maximizado.
- O script cria um perfil isolado em chrome-profile/ para preservar login/cookies sem afetar seu Chrome padrão.
