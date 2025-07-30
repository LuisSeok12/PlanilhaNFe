import os
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import csv
import re
import pdfplumber
from datetime import datetime


def criar_nome_seguro(nome):
    for c in '<>:"/\\|?*': nome = nome.replace(c, '')
    return nome[:15]


def formatar_data(data_iso):
    try:
        return datetime.fromisoformat(data_iso).strftime("%d/%m/%Y")
    except:
        return "Data Inválida"


def extrair_informacoes_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        texto = pdf.pages[0].extract_text() or ''
    linhas = texto.split('\n')
    nome = None; nf = None; serie = None

    p1 = re.compile(r'RECEBEMOS\s+D[AE]\s+(.*?)\s+OS\s+PRODUTOS', re.I)
    for l in linhas:
        m = p1.search(l)
        if m:
            nome = m.group(1).strip(); break
    if not nome:
        for l in linhas:
            if 'RAIZEN POWER' in l.upper():
                nome = 'RAIZEN POWER COMERCIALIZADORA DE EN' ; break
    # NF
    for l in linhas:
        m = re.search(r'(?:N[º°]?\.?|No\.?)\s*([\d\.]+)', l)
        if m and m.group(1).replace('.','').isdigit():
            nf = m.group(1).replace('.',''); break
    # Série
    for l in linhas:
        m = re.search(r'S[ÉE]RIE[:\s]*(\d+)', l, re.I)
        if m: serie = m.group(1); break
    if nome: nome = re.sub(r'^RECEBEMOS\s+D[AE]\s+', '', nome, flags=re.I)
    return nome, nf, serie


def processar_arquivo_xml(caminho, pasta, out):
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()
    except:
        return
    ns = {'ns':'http://www.portalfiscal.inf.br/nfe'}
    num = root.find('ns:NFe/ns:infNFe/ns:ide/ns:nNF', ns)
    ser = root.find('ns:NFe/ns:infNFe/ns:ide/ns:serie', ns)
    emi = root.find('ns:NFe/ns:infNFe/ns:ide/ns:dhEmi', ns)
    emit = root.find('ns:NFe/ns:infNFe/ns:emit/ns:xNome', ns)
    prot = root.find('ns:protNFe/ns:infProt/ns:nProt', ns)
    if num is not None and ser is not None and emit is not None:
        nf_c = f"{num.text}-{ser.text}"
        out.append([emit.text, nf_c, root.find('ns:NFe/ns:infNFe', ns).attrib.get('Id','')[3:], prot.text if prot else '', formatar_data(emi.text if emi else '')])
        novo = f"{criar_nome_seguro(emit.text)} NF {nf_c}.xml"
        os.rename(caminho, os.path.join(pasta, novo))


def processar_arquivo_pdf(caminho, pasta):
    nome, nf, serie = extrair_informacoes_pdf(caminho)
    if nome and nf:
        base = nome.split()[0]
        novo = f"{base} NF {nf}{('-'+serie) if serie else ''}.pdf"
        os.rename(caminho, os.path.join(pasta, novo))


def run_leitor_pdf_xml():
    root = tk.Tk(); root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione pasta com XML/PDF")
    if not pasta: return
    resultados = []
    for f in os.listdir(pasta):
        p = os.path.join(pasta, f)
        if f.lower().endswith('.xml'):
            processar_arquivo_xml(p, pasta, resultados)
        elif f.lower().endswith('.pdf'):
            processar_arquivo_pdf(p, pasta)
    csv_p = os.path.join(pasta, 'notas.csv')
    with open(csv_p, 'w', newline='', encoding='utf-8') as f:
        w=csv.writer(f); w.writerow(["Emissor","NF Completa","Chave","Protocolo","Data de Emissão"]); w.writerows(resultados)
    messagebox.showinfo("Sucesso", f"Dados salvos em {csv_p}")