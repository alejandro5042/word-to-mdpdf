# Converts an Word doc to Markdown/PDF

Such as specs like [IVI 3.1 Architecture](https://www.ivifoundation.org/downloads/Architecture%20Specifications/IVI-3.1_Architecture_2019-08-12.pdf).

## Pre-requisites

- [pandoc](https://pandoc.org/) (to convert Word to Markdown)
    - e.g. `scoop install pandoc`
- [wkhtmltopdf](https://wkhtmltopdf.org/) (to convert Markdown to PDF)
    - e.g. `scoop install wkhtmltopdf`
- [Docker](https://www.docker.com/) (to serve the demo Jekyll site)

## Sample Usage

```powershell
.\scripts\Convert-WordToMD.ps1 .\originals\IVI-3.1_Architecture_2022-12-19.docx
.\scripts\serve.ps1 .\out\IVI-3.1_Architecture_2022-12-19\
```

## Getting Help

Talk to Alejandro Barreto.
