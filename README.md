# JDOfx

Este programa recebe como parâmetro o caminho de um arquivo com extensão ofx e converte os dados de transações em uma planilha no formato Microsft Excel


## Autores

- [@duartejd](https://github.com/DuarteJD)


### Dependencia para criação de um executável desta aplicação

ofxparse
>https://pypi.org/project/ofxparse/

openpyxl
>https://openpyxl.readthedocs.io/en/stable/

pyinstaller  
>https://pyinstaller.org/en/stable/

```bash
  pip install ofxparse openpyxl pyinstaller
```

### Parâmetros necessários para utilização do programa
**Parâmetro 01**: string contendo o path do arquivo .ofx  
**Parâmetro 02**: caminho completo do arquivo que será criado com o resultado do processamento

### Exemplo de utilização
Parâmetro 01: c:\extrato.ofx
Parâmetro 02: c:\extrato.xlsx
Comando: **python main.py c:\extrato.ofx c:\extrato.xlsx** 

### Como criar um executável desta aplicação
No terminal digite o comando:  
    pyinstaller main.py -F -n JDOfx  

<a name="my-custom-anchor-point"></a>
    -F: Cria um único arquivo executável com todas as dependências incluídas  
    -n: nome do arquivo executável desejado. (Padrão: nome da rotina principal, neste caso seria: main.exe)  