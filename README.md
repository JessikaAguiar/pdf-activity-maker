# pdf-activity-maker



## Installar as dependencias

``` 
pip install -r requirements.txt
```


## Pré-requisitos

* Python 3.13 já instalado

### Tesseract OCR instalado e configurado no PATH

💡 Baixe o instalador do Tesseract: https://github.com/tesseract-ocr/tesseract/releases

Durante a instalação:

Marque a opção "Add to PATH"

Selecione o idioma Portuguese (por)


1. Acesse a pasta onde o Tesseract foi instalado:
makefile
Copiar
Editar
C:\Program Files\Tesseract-OCR\tessdata
Veja se existe um arquivo chamado por.traineddata.
Se não existir, siga abaixo 👇

2. 🔽 Baixar o idioma português (por.traineddata)
Baixe o arquivo diretamente aqui:

📄 https://github.com/tesseract-ocr/tessdata/blob/main/por.traineddata