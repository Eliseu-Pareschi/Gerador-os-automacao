import base64
from datetime import datetime
import os
import re
import time
import traceback
from flask import Flask, jsonify, render_template, request
import pythoncom
import requests
import win32com.client as win32
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env local
load_dotenv()

app = Flask(__name__)

# Puxa dados sensíveis e configurações de forma segura do .env
FIREBASE_URL = os.getenv("FIREBASE_URL")
EMAIL_CC_CONFIG = os.getenv("EMAIL_CC", "")
NOME_AUTOR = os.getenv("MEU_NOME", "Equipe de Suporte")
TELEFONE_AUTOR = os.getenv("MEU_TELEFONE", "")
EMAIL_AUTOR = os.getenv("MEU_EMAIL", "")

CAMINHO_MODELO_XLSM = r"C:\Users\e.lopespares\Desktop\ASSISTENTE\modelo OS.xlsm"
CAMINHO_IMAGEM_ASSINATURA = os.path.abspath(
    os.path.join("templates", "assinatura.png")
)

# Pasta dedicada para salvar os PDFs das Ordens de Serviço de forma organizada
PASTA_OS = os.path.abspath("ordens_servico")
os.makedirs(PASTA_OS, exist_ok=True)

# Imagem PNG transparente em Base64 para fallback quando não houver foto
DUMMY_PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="


def limpar_string(texto):
    """Remove caracteres especiais e espaços, mantendo apenas letras e números em caixa alta."""
    if not texto:
        return ""
    return re.sub(r"[^A-Z0-9]", "", texto.upper())


def limpar_nome_arquivo(texto):
    """Remove caracteres inválidos para nomes de arquivos no Windows (ex: / \ : * ? \" < > |)."""
    if not texto:
        return ""
    return re.sub(r'[\\/*?:"<>|]', "", str(texto)).strip()


def criar_imagem_dummy(caminho_destino):
    """Cria uma imagem PNG vazia para evitar erros no Excel caso o usuário não envie foto."""
    try:
        with open(caminho_destino, "wb") as fh:
            fh.write(base64.b64decode(DUMMY_PNG_B64))
    except Exception:
        pass


@app.route("/")
def index():
    return render_template("interface.html")


@app.route("/obter-dados-base/<path:nome_base>", methods=["GET"])
def obter_dados_base(nome_base):
    try:
        resp = requests.get(FIREBASE_URL, timeout=10)
        todas_bases = resp.json()

        if not todas_bases:
            return jsonify({"erro": "Nenhuma base encontrada"}), 404

        chave_recebida_limpa = limpar_string(nome_base)
        dados_base = {}

        for k, v in todas_bases.items():
            if limpar_string(k) == chave_recebida_limpa:
                if isinstance(v, dict):
                    dados_base = v
                break

        if not dados_base:
            return jsonify({"erro": "Base não encontrada"}), 404

        return jsonify(
            {
                "endereco": dados_base.get("endereco", ""),
                "coordenador": dados_base.get("coordenador", ""),
                "email": dados_base.get("email", ""),
            }
        )
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/gerar-os", methods=["POST"])
def gerar_os():
    caminho_antes = ""
    caminho_depois = ""
    excel = None
    wb = None

    try:
        # Inicializa as threads do COM do Windows para evitar erros no Flask
        pythoncom.CoInitialize()

        # Coleta de dados do Formulário
        num_os = request.form.get("numOS", "").strip()
        base_key_recebida = request.form.get("baseKey", "").strip()
        descricao = request.form.get("descricao", "").strip()
        atribuido_para = request.form.get(
            "atribuidoPara", "GUSTAVO/ MARCIO/ GABRIEL"
        )
        trabalhos_efetuados = request.form.get("trabalhosEfetuados", "")
        area_servico = request.form.get("area", "ALMOXARIFADO")

        # Processamento das imagens enviadas
        timestamp = int(time.time())
        foto_antes = request.files.get("fotoAntes")
        foto_depois = request.files.get("fotoDepois")

        caminho_antes = os.path.abspath(f"temp_antes_{timestamp}.png")
        if foto_antes and foto_antes.filename != "":
            foto_antes.save(caminho_antes)
        else:
            criar_imagem_dummy(caminho_antes)

        caminho_depois = os.path.abspath(f"temp_depois_{timestamp}.png")
        if foto_depois and foto_depois.filename != "":
            foto_depois.save(caminho_depois)
        else:
            criar_imagem_dummy(caminho_depois)

        # Consulta os dados correspondentes no Firebase Realtime Database de forma segura
        resp = requests.get(FIREBASE_URL, timeout=10)
        todas_bases = resp.json() or {}

        chave_receb_limpa = limpar_string(base_key_recebida)
        dados_base = {}
        base_key_real = base_key_recebida

        for k, v in todas_bases.items():
            if limpar_string(k) == chave_receb_limpa:
                base_key_real = k
                if isinstance(v, dict):
                    dados_base = v
                break

        if not isinstance(dados_base, dict):
            dados_base = {}

        endereco = dados_base.get("endereco", "VAZIO")
        coordenador = dados_base.get("coordenador", "Equipe")
        email_destinatario = dados_base.get("email", "VAZIO")
        data_atual = datetime.now().strftime("%d/%m/%Y")

        # Execução do Excel em plano de fundo via COM
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(CAMINHO_MODELO_XLSM)

        # Formatação do nome do arquivo PDF e salvamento dentro da pasta "ordens_servico"
        nome_arquivo_cru = f"{num_os}, {base_key_real}, {descricao}"
        nome_arquivo_seguro = limpar_nome_arquivo(nome_arquivo_cru)
        nome_pdf_saida = os.path.join(PASTA_OS, f"{nome_arquivo_seguro}.pdf")

        # Chamada da Macro interna do arquivo .xlsm
        nome_macro = f"'{wb.Name}'!GerarOS_PDF"

        excel.Run(
            nome_macro,
            num_os,
            base_key_real,
            endereco,
            coordenador,
            email_destinatario,
            descricao,
            atribuido_para,
            trabalhos_efetuados,
            area_servico,
            data_atual,
            caminho_antes,
            caminho_depois,
            nome_pdf_saida,
        )

        # Fecha a planilha sem salvar alterações no modelo base
        wb.Close(SaveChanges=False)
        wb = None
        excel.Quit()
        excel = None

        # Criação da mensagem de e-mail e anexo no Microsoft Outlook
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = email_destinatario
        mail.CC = EMAIL_CC_CONFIG
        
        # Assunto formatado em caixa alta com os parâmetros
        mail.Subject = f"ORDEM DE SERVIÇO, {num_os}, {base_key_real}, {descricao}"
        
        # Corpo do e-mail estruturado utilizando variáveis protegidas
        mail.HTMLBody = f"""
        <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px; color: #000000;">
            <p>Boa tarde {coordenador}</p>
            <br>
            <p>Segue a ordem de serviço referente ao reparo realizado na base, para arquivo</p>
            <br>
            <p>Atenciosamente.</p>
            <br>
            <p>
                <b>{NOME_AUTOR}</b><br>
                INFRAESTRUTURA<br>
                Carso Instalações do Brasil<br>
                {TELEFONE_AUTOR} / {EMAIL_AUTOR}
            </p>
            <br>
            <img src="cid:logo_carso" style="max-width: 250px; height: auto;">
        </body>
        </html>
        """
        
        # Anexa a logo corretamente usando o ID inline
        if os.path.exists(CAMINHO_IMAGEM_ASSINATURA):
            attachment = mail.Attachments.Add(CAMINHO_IMAGEM_ASSINATURA)
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_carso")

        # Anexa o PDF gerado diretamente da pasta organizada
        if os.path.exists(nome_pdf_saida):
            mail.Attachments.Add(nome_pdf_saida)
            
        mail.Display(False)

        return render_template("sucesso.html", num_os=num_os)

    except Exception:
        erro_completo = traceback.format_exc()
        return (
            f"<h3>Ocorreu um erro ao processar a OS:</h3><pre>{erro_completo}</pre>",
            500,
        )

    finally:
        if excel:
            try:
                if wb:
                    wb.Close(SaveChanges=False)
                excel.Quit()
            except Exception:
                pass

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        if caminho_antes and os.path.exists(caminho_antes):
            try:
                os.remove(caminho_antes)
            except Exception:
                pass

        if caminho_depois and os.path.exists(caminho_depois):
            try:
                os.remove(caminho_depois)
            except Exception:
                pass


if __name__ == "__main__":
    app.run(port=5000, debug=True, threaded=False)