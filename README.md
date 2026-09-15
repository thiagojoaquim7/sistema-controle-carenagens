# Sistema de Controle de Carenagens

Sistema web desenvolvido para digitalizar, organizar e acompanhar o processo operacional de recebimento, conferência e entrada de carenagens.

## Sobre o projeto

O sistema foi desenvolvido com o objetivo de centralizar o controle das etapas relacionadas ao recebimento de carenagens, reduzindo controles manuais e aumentando a rastreabilidade das operações.

A aplicação permite registrar entregas, realizar conferências, identificar divergências, anexar documentos fiscais e acompanhar o andamento de cada recebimento.

## Principais funcionalidades

- Autenticação de usuários
- Controle de acesso por função
- Cadastro de novas entregas
- Registro de recebimentos
- Cadastro e consulta de peças
- Preenchimento automático de descrições através da base de referências
- Conferência de quantidades
- Identificação de divergências
- Registro de avarias
- Controle da adequação da embalagem
- Anexação de NF-e
- Visualização da NF-e diretamente no sistema
- Conferência da NF-e
- Aprovação ou rejeição da NF-e
- Controle das entradas pendentes
- Histórico das movimentações
- Gerenciamento de usuários
- Alteração e redefinição de senhas
- Integração com banco de dados SQLite
- Integração com arquivo Excel para consulta de referências

## Fluxo do processo

```text
Nova Entrega
     |
     v
Recebimento
     |
     v
Conferência Física
     |
     v
Anexação da NF-e
     |
     v
Conferência da NF-e
     |
     +------------------+
     |                  |
     v                  v
Aprovada            Rejeitada
     |
     v
Entrada Realizada
     |
     v
Finalização
