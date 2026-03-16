# Certificado Automation Python 🚀

Sistema desenvolvido para automatizar a extração de dados de ordens de serviço em sistemas web e o preenchimento automático de certificados de garantia em Excel.

## 📋 Descrição do Projeto
Este projeto resolve um problema real de produtividade em óticas, eliminando a necessidade de redigitação manual de prescrições e dados de pacientes. Ele integra uma interface web via JavaScript (Bookmarklet) com um servidor local em Python para manipulação de arquivos .xls.

## 🛠️ Tecnologias Utilizadas
- **Python**: Linguagem principal para o processamento de dados.
- **Flask**: Micro-framework para criação da API de recebimento de dados.
- **xlrd / xlwt / xlutils**: Bibliotecas para leitura e escrita em arquivos Excel legados.
- **JavaScript**: Script de front-end para extração de dados via DOM.

## ⚙️ Funcionalidades
- **Extração Inteligente**: Captura automática de prescrições (Esférico, Cilíndrico, Eixo, DNP, etc).
- **Tratamento de Dados**: Conversão de nomenclaturas técnicas para nomes comerciais de certificados.
- **Lógica de Produtos**: Identificação automática de tratamentos fotossensíveis (TGNS/FOTO CINZA).
- **Saída Estruturada**: Gravação de IDs de pedidos como valores numéricos para compatibilidade com sistemas de etiquetas.

## 🚀 Como Executar
1. Instale as dependências: `pip install flask flask-cors xlrd xlwt xlutils`
2. Execute o servidor: `python certificado_automation.py`
3. Utilize o código contido em `bookmarklet.js` como um favorito no seu navegador.
4. Clique no favorito dentro da página do pedido para processar os dados.

---
Desenvolvido por Gabriel Sales
