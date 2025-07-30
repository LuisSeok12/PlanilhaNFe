import sys
from gui import main_menu
from planilhar import run_planilhar_notas
from leitor import run_leitor_pdf_xml

if __name__ == '__main__':
    # Verifica se Pillow está instalado
    try:
        import PIL
    except ImportError:
        print("Este programa requer a biblioteca PIL/Pillow.")
        print("Instale com: pip install pillow")
        sys.exit(1)

    if len(sys.argv) > 1:
        if sys.argv[1] == "planilhar":
            run_planilhar_notas()
        elif sys.argv[1] == "leitor":
            run_leitor_pdf_xml()
        else:
            print(f"Argumento desconhecido: {sys.argv[1]}")
    else:
        main_menu()
