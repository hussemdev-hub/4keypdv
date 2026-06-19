# 4keypdv

Aplicação desktop desenvolvida em Python para automação de operações no MaxPOS, incluindo alteração de senhas PDV, envio de carga, processamento em lote, histórico de usuários e integração via Selenium.

## Instalação (Linux)

```bash
sudo apt update
sudo apt install python3 python3-pip
pip install -r requirements.txt
python3 main.py
```

## Configuração

Na primeira execução, o arquivo de configuração é criado automaticamente em:

```
~/.local/share/4keypdv/config.ini
```

Edite esse arquivo para apontar para as URLs e filial do seu ambiente:

```ini
[PDV]
PDV_URL = http://seu-servidor/maxipos_backoffice/app
MAESTRO_URL = https://login.seu-servidor.com.br
FILIAL = 42
HEADLESS = True
```

## Dados salvos

```
~/.local/share/4keypdv/
├── config.ini
├── 4keypdv_history.xlsx
├── 4keypdv_credentials.xlsx
└── logs/
    └── 4keypdv.log
```
