import socket

from pptx import Presentation
import os
from flask import Flask, send_from_directory

from flask import request

gerarArquivo = Flask("name")


# pip install python-pptx


@gerarArquivo.route("/")
def gerar():
    try:
        return "<p>sucesso AO ENTrar</p>"
    except:
        return "<p>erro</p>"

@gerarArquivo.route("/gerar", methods=['POST'])
def gerar_slides():
    # metodo para gerar e salvar o slides com a letra
    # recebendo os dados via POST
    textos = request.form.get('textos')
    titulo = request.form.get('titulo')
    modelo = request.form.get('modelo')
    try:
        # dividindo o texto em um lista apartir da seguinte separacao
        textoDividido = textos.split("</p>")
        # removendo o ultimo index da lista pois o mesmo fica vazio
        textoDividido.pop(-1)
        if 'geral' in modelo:
            caminho = "modelos_slides/modelo_geral.pptx"
            prs = Presentation(caminho)
        else:
            caminho = "modelos_slides/modelo_geracao_fire.pptx"
            prs = Presentation(caminho)
        blank_slide_layout = prs.slide_layouts[0]
        for estrofe in textoDividido:  # pegando cada item da lista
            dividir_estrofe = estrofe.split("<br>")
            # divindo cada item em uma nova lista
            for verso in dividir_estrofe:  # pegando cada item da nova lista gerada
                # adicionando um slide para cada item da lista
                slide = prs.slides.add_slide(blank_slide_layout)
                title = slide.shapes.title
                subtitulo = slide.placeholders[1]
                title.text = titulo
                subtitulo.text = verso

        prs.save(titulo + ".pptx")  # salvando o arquivo
        return "<p>sucesso</p>"
    except:
        return "<p>erro</p>"


def obterIP():
    # metodo para obter ip da maquina
    try:
        # obtendo ip
        ip = socket.gethostbyname(socket.gethostname())
        print(ip)
        # salvando o arquivo
        return ip
    except:
        return "<p>erro</p>"


@gerarArquivo.route("/baixarArquivo/<nome_arquivo>", methods=['GET'])
def baixar_arquivo(nome_arquivo):
    try:
        caminho_absoluto_arquivo_python = os.path.abspath(__file__)
        diretorio_src = os.path.dirname(caminho_absoluto_arquivo_python)
        print(diretorio_src)
        arquivo = nome_arquivo + ".pptx"
        return send_from_directory(diretorio_src,
                                   arquivo, as_attachment=True)
    except:
        return "<p>erro</p>"


@gerarArquivo.route("/excluirArquivo", methods=['POST'])
def excluir_arquivo():
    # metodo para excluir o arquivo gerado
    arquivo = request.form.get('arquivo')
    nome_arquivo = arquivo + ".pptx"
    caminho_absoluto_arquivo_python = os.path.abspath(__file__)
    diretorio_src = os.path.dirname(caminho_absoluto_arquivo_python)
    diretorio = os.listdir(diretorio_src)
    try:
        for file in diretorio:
            if file == nome_arquivo:
                os.remove(file)
        return "<p>sucesso ao excluir</p>"
    except:
        return "<p>erro ao excluir</p>"


if __name__ == '__main__':
    gerarArquivo.run(host=obterIP(), debug=True)
