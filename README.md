# 📄 Utilitário de Automação e Processamento de PDFs para RH

> Um utilitário de desktop leve, desenvolvido em Python com CustomTkinter, criado para automatizar a separação de contracheques, folhas de ponto e o cruzamento inteligente de comprovativos bancários para o departamento de Recursos Humanos.

---

## 🚀 Principais Funcionalidades

* **Automação baseada em Excel:** Lê listas de colaboradores diretamente de tabelas Excel (`.xlsx`) ou CSV para filtrar e segmentar lotes de PDFs de forma instantânea.
* **Modo de Emergência Híbrido (OCR Integrado):** Se o PDF original for uma imagem escaneada (onde o texto não é selecionável), o sistema aciona automaticamente o motor **Tesseract OCR** para transcrever o documento em tempo de execução.
* **Resolução Visual de Conflitos (Adiantamentos vs. Salários):** Quando o sistema deteta ambiguidades em valores de adiantamentos, pausa a execução de forma inteligente e abre uma pré-visualização em alta resolução da página do PDF (`PyMuPDF`) para validação humana em tela cheia.
* **Arquitetura Portátil (Portable App):** Desenhado para correr de forma autónoma sem necessidade de instalações complexas no sistema operativo.

---

## 🛠️ Tecnologias Utilizadas

* **Python** (Linguagem principal)
* **CustomTkinter** (Interface gráfica moderna e elegante)
* **PyPDF & PyMuPDF (fitz)** (Manipulação e renderização avançada de PDFs)
* **Tesseract-OCR** (Reconhecimento ótico de caracteres para documentos escaneados)
* **OpenPyXL** (Leitura e processamento de folhas de cálculo Excel)

---

## 📦 Como Executar o Projeto

1. Certifique-se de que tem o Python instalado no seu computador.
2. Instale as dependências necessárias executando o seguinte comando no terminal:
   ```bash
   pip install customtkinter pypdf pymupdf pillow openpyxl
