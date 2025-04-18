# Ativador de Windows KMS

Este programa permite ativar o Windows e instalar o Microsoft Office de forma simples e rápida.

## Funcionalidades

- Ativação automática do Windows 11, 10 e 8
- Instalação de diferentes versões do Office (2019 e 2021)
- Suporte para arquiteturas x86 e x64
- Interface gráfica moderna e intuitiva

## Requisitos

- Python 3.8 ou superior
- PyQt6
- PyInstaller (apenas para compilação)

## Como compilar

1. Instale as dependências:
```bash
pip install -r requirements.txt
```

2. Execute o script de compilação:
```bash
python build.py
```

3. O executável será gerado na pasta `dist`

## Como usar

1. Execute o programa como administrador
2. Selecione a versão do Office desejada (ou escolha não instalar)
3. Clique em "Ativar Windows"
4. Aguarde o processo ser concluído

## Observações

- O programa precisa ser executado como administrador para funcionar corretamente
- Certifique-se de ter conexão com a internet para a ativação
- O servidor KMS usado é: carlosandradepy-39464.portmap.host:44258

## Suporte

Para mais informações ou suporte, entre em contato através do GitHub. 