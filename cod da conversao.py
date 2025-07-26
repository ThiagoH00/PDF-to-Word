from pdf2docx import Converter

def converter_pdf_para_word_local(caminho_pdf, caminho_word):
    """
    Converte um arquivo PDF para um arquivo DOCX localmente.

    :param caminho_pdf: O caminho completo para o arquivo PDF de entrada.
    :param caminho_word: O caminho completo para salvar o arquivo DOCX de saída.
    """
    try:
        # Cria um objeto conversor
        cv = Converter(caminho_pdf)
        
        # Converte o PDF para DOCX
        # start=0, end=None significa converter todas as páginas
        cv.convert(caminho_word, start=0, end=None)  # type: ignore
        
        # Fecha o objeto para liberar o arquivo
        cv.close()
        
        print(f"Sucesso! Arquivo '{caminho_pdf}' convertido para '{caminho_word}'.")
        
    except FileNotFoundError:
        print(f"Erro: O arquivo de entrada '{caminho_pdf}' não foi encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro durante a conversão: {e}")

if __name__ == "__main__":
    # --- ATENÇÃO: Altere os caminhos dos arquivos aqui ---
    arquivo_pdf_entrada = "exemplo.pdf"  # Coloque o nome do seu PDF aqui
    arquivo_word_saida = "resultado.docx" # Nome do arquivo Word que será criado

    converter_pdf_para_word_local(arquivo_pdf_entrada, arquivo_word_saida)
