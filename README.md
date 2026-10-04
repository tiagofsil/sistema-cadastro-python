# Sistema de Cadastro em Python

Programa de terminal que cadastra pessoas e salva os dados em uma planilha Excel (.xlsx).

## Funcionalidades
- Validação de e-mail, data de nascimento e telefone
- Bloqueio de e-mails duplicados
- Formatação automática de nomes e cidades
- Cria a planilha ou acrescenta na existente

## Como rodar
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 cadastro.py
```

## Tecnologias
Python, openpyxl, regex
