# 🧾 Auditoria NCM - Validador Fiscal

Um sistema focado em automação comercial desenvolvido para prevenir rejeições de NFC-e na SEFAZ-SP. A aplicação identifica e valida códigos NCM (Nomenclatura Comum do Mercosul) vencidos ou desatualizados diretamente no banco de dados de frente de caixa do cliente.

## 🎯 O Problema
No dia a dia do varejo, a rejeição de notas fiscais (NFC-e) devido a cadastros de produtos com NCMs inativos causa paradas no caixa, gerando transtornos para o cliente e gargalos para o suporte técnico.

## 💡 A Solução
Esta aplicação atua de forma preventiva. Ela se conecta ao banco de dados do sistema PDV (Firebird), extrai os NCMs cadastrados e os valida em lote consumindo a **Brasil API**. Isso permite que a equipe ou o próprio lojista corrija as tributações antes que a venda seja bloqueada no momento da emissão.

## 🚀 Funcionalidades
- **Conexão Nativa:** Integração segura com bancos de dados Firebird utilizando a dll `FirebirdSql.Data.FirebirdClient`.
- **Auditoria Automatizada:** Varredura e leitura dos códigos NCM vinculados ao cadastro de produtos.
- **Validação em Tempo Real:** Consumo de endpoints da Brasil API para checar a vigência fiscal dos códigos.
- **Prevenção de Erros:** Redução drástica de chamados de suporte relacionados a rejeições da SEFAZ-SP.

## 🛠️ Tecnologias Utilizadas
- **Linguagem:** PowerShell
- **Banco de Dados:** Firebird SQL
- **Dependências:** `FirebirdSql.Data.FirebirdClient`
- **APIs Externas:** [Brasil API](https://brasilapi.com.br/)

## ⚙️ Como Executar o Projeto

### Pré-requisitos
- [Ex: .NET SDK] instalado na máquina.
- Servidor Firebird em execução.
- Banco de dados do cliente (arquivo `.fdb`) acessível.
- Conexão com a internet para consumo da Brasil API.
 
<img width="1365" height="738" alt="consulta ncm" src="https://github.com/user-attachments/assets/756c4d53-b2db-40f4-9251-1856c69b49fb" />

