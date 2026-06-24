# Automação de Declaração de Conteúdo (Logística TI)

### O Problema

No dia a dia de Field Service, o preenchimento manual de declarações de conteúdo para envio de equipamentos via Sedex era um processo repetitivo e sujeito a erros de digitação, consumindo tempo precioso da equipe de suporte.

### A Solução

Desenvolvi um script em Python que extrai dados brutos de chamados técnicos que preenchemos e depois preenche automaticamente um template .docx, convertendo-o em PDF pronto para impressão via LibreOffice.

### Tecnologias Utilizadas

    Python 3.x

    Library python-docx para manipulação de documentos Word.

    LibreOffice (Soffice) para conversão headless em PDF.

    Datetime & Subprocess para automação de sistema.

### Como Usar

    Instale as dependências: pip install -r requirements.txt.

    Certifique-se de ter o LibreOffice instalado no caminho configurado no script.

    Coloque o texto do chamado na variável texto_bruto.

    Execute: python src/preencherDeclaracao.py
