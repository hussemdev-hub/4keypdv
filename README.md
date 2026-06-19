# 4keypdv

Automação de alteração de senha PDV e envio de carga via Selenium + Tkinter.

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