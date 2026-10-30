"""Gera o documento .docx de levantamento de bens patrimoniais.

O layout segue o modelo usado pelo setor de patrimônio: título com o local,
tabela DESCRIÇÃO/PATRIMÔNIO, linha com o responsável e, em seguida, uma foto
de cada bem por página, na mesma ordem da tabela.

Uso:
    python3 gerar_levantamento.py levantamentos/<pasta>/levantamento.json
"""
import argparse
import json
import logging
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Cm, Pt
from PIL import Image

LARGURA_PAGINA_A4 = Cm(21)
ALTURA_PAGINA_A4 = Cm(29.7)
MARGEM = Cm(2)
FONTE_PADRAO = 'Calibri'
TAMANHO_FONTE = Pt(11)

logger = logging.getLogger(__name__)


def carregar_levantamento(caminho_json):
    """Lê o JSON do levantamento e valida os campos obrigatórios.

    Args:
        caminho_json (Path): arquivo com as chaves ``local``, ``responsavel`` e ``itens``.

    Returns:
        dict: dados do levantamento.

    Raises:
        ValueError: se faltar algum campo ou alguma foto referenciada não existir.
    """
    dados = json.loads(caminho_json.read_text(encoding='utf-8'))
    for campo in ('local', 'responsavel', 'itens'):
        if not dados.get(campo):
            raise ValueError(f'Campo obrigatório ausente no levantamento: {campo}')

    for posicao, item in enumerate(dados['itens'], start=1):
        if not item.get('descricao') or not item.get('patrimonio'):
            raise ValueError(f'Item {posicao} sem descrição ou patrimônio')
        if item.get('foto') and not (caminho_json.parent / item['foto']).is_file():
            raise ValueError(f'Foto do item {posicao} não encontrada: {item["foto"]}')
    return dados


def configurar_documento():
    """Cria um documento A4 retrato com a fonte padrão do modelo."""
    documento = Document()
    secao = documento.sections[0]
    secao.orientation = WD_ORIENT.PORTRAIT
    secao.page_width = LARGURA_PAGINA_A4
    secao.page_height = ALTURA_PAGINA_A4
    secao.left_margin = secao.right_margin = MARGEM
    secao.top_margin = secao.bottom_margin = MARGEM

    estilo = documento.styles['Normal']
    estilo.font.name = FONTE_PADRAO
    estilo.font.size = TAMANHO_FONTE
    return documento


def adicionar_titulo(documento, local):
    paragrafo = documento.add_paragraph()
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragrafo.paragraph_format.space_after = Pt(36)
    paragrafo.add_run(f'LEVANTAMENTO DE BENS PATRIMONIAIS {local}'.upper())


def adicionar_tabela_bens(documento, itens):
    tabela = documento.add_table(rows=1, cols=2)
    tabela.style = 'Table Grid'
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER

    for celula, titulo in zip(tabela.rows[0].cells, ('DESCRIÇÃO', 'PATRIMÔNIO')):
        _escrever_celula_centralizada(celula, titulo, negrito=True)

    for item in itens:
        celula_descricao, celula_patrimonio = tabela.add_row().cells
        _escrever_celula_centralizada(celula_descricao, item['descricao'].upper())
        _escrever_celula_centralizada(celula_patrimonio, item['patrimonio'].upper())


def _escrever_celula_centralizada(celula, texto, negrito=False):
    paragrafo = celula.paragraphs[0]
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragrafo.paragraph_format.space_after = Pt(0)
    paragrafo.add_run(texto).bold = negrito


def adicionar_responsavel(documento, responsavel):
    paragrafo = documento.add_paragraph()
    paragrafo.paragraph_format.space_before = Pt(48)
    paragrafo.add_run(f'RELAÇÃO DE BENS FEITA POR {responsavel}.'.upper())


def adicionar_fotos(documento, itens, pasta_base):
    """Insere uma foto por página, redimensionada para caber na área útil."""
    largura_util = LARGURA_PAGINA_A4 - 2 * MARGEM
    altura_util = ALTURA_PAGINA_A4 - 2 * MARGEM - Cm(1)

    for item in itens:
        if not item.get('foto'):
            continue
        caminho_foto = pasta_base / item['foto']
        with Image.open(caminho_foto) as imagem:
            largura_px, altura_px = imagem.size

        escala = min(largura_util / largura_px, altura_util / altura_px)
        paragrafo = documento.add_paragraph()
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragrafo.add_run()
        run.add_break(WD_BREAK.PAGE)
        run.add_picture(str(caminho_foto), width=int(largura_px * escala), height=int(altura_px * escala))


def gerar_documento(caminho_json, caminho_saida):
    """Monta e salva o .docx do levantamento descrito em ``caminho_json``."""
    dados = carregar_levantamento(caminho_json)
    documento = configurar_documento()
    adicionar_titulo(documento, dados['local'])
    adicionar_tabela_bens(documento, dados['itens'])
    adicionar_responsavel(documento, dados['responsavel'])
    adicionar_fotos(documento, dados['itens'], caminho_json.parent)
    documento.save(caminho_saida)
    return caminho_saida


def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Gera o .docx de levantamento patrimonial.')
    parser.add_argument('levantamento', type=Path, help='Caminho do levantamento.json')
    parser.add_argument('--saida', type=Path, help='Arquivo .docx de saída (padrão: ao lado do JSON)')
    argumentos = parser.parse_args()

    saida = argumentos.saida or argumentos.levantamento.with_name('LEVANTAMENTO_PATRIMONIO.docx')
    try:
        gerar_documento(argumentos.levantamento, saida)
    except (OSError, ValueError, json.JSONDecodeError) as erro:
        logger.error('Não foi possível gerar o levantamento: %s', erro)
        sys.exit(1)
    logger.info('Documento gerado em %s', saida)


if __name__ == '__main__':
    main()
