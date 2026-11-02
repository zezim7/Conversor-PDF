# Conversor PDF

Programa para Windows que transforma quase qualquer arquivo em PDF — e junta tudo em um único arquivo, na ordem que você quiser.

## O que converte

- **Imagens**: JPG, PNG, WEBP, HEIC, BMP, GIF, TIFF (várias páginas), ICO, prints colados (Ctrl+V)
  - 1, 2, 4 ou 6 imagens por página, girar, encaixar ou preencher, papel A4/Carta/Ofício
  - **Modo scanner**: deixa o fundo de fotos de documentos branco e tira sombras
- **PDF**: junta vários PDFs em um só
- **Word, Excel, PowerPoint** (.docx, .xlsx, .pptx): usa o Microsoft Office ou o LibreOffice, se instalados; senão, um conversor interno (layout aproximado)
- **Textos**: TXT, Markdown, CSV (vira tabela), JSON, XML, código
- **Web**: HTML, MHT, SVG ou o endereço de um site
- **E-mail**: arquivos .eml

## Requisitos

- Windows 10 ou 11 (usa o .NET Framework 4.x e o Microsoft Edge, que já vêm no Windows)
- Nenhuma biblioteca externa

## Como compilar

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1
```

O executável sai em `bin\Conversor PDF.exe`. O script usa o compilador C# que já vem no Windows (`csc.exe` do .NET Framework).

## Estrutura

| Arquivo | Conteúdo |
|---|---|
| `src/Program.cs` | Interface (janela, lista, prévia, geração) |
| `src/Imaging.cs` | Leitura de imagens, miniaturas, modo scanner, compressão |
| `src/PdfOut.cs` | Escrita do PDF e diagramação das imagens nas páginas |
| `src/PdfIn.cs` | Leitor de PDF (para juntar PDFs existentes) |
| `src/Converters.cs` | Conversão via Edge/Office/LibreOffice, texto, Markdown, CSV, e-mail |
| `src/OfficeBasico.cs` | Conversor interno de .docx, .xlsx e .pptx para HTML |
| `build.ps1` | Script de compilação (gera também o ícone) |
