# Automação de Levantamento Patrimonial

Utilitário em Python para gerar relatórios `.docx` padronizados de inventário/patrimônio a partir de arquivos JSON e fotos dos bens.

## O que resolve

- Padroniza tabelas de descrição e número de patrimônio em A4;
- Valida campos obrigatórios e arquivos de imagem antes de gerar o documento;
- Anexa automaticamente uma foto por página na mesma ordem da tabela.

## Como usar

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python3 gerar_levantamento.py exemplos/levantamento_exemplo.json
```
