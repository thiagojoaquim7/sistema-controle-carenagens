# Sistema de Controle de Carenagens

Sistema web desenvolvido para digitalizar e organizar o processo de recebimento, conferência e acompanhamento de carenagens.

## Sobre o projeto

O sistema foi desenvolvido para centralizar o controle operacional das entradas de carenagens, permitindo acompanhar cada etapa do processo desde o recebimento até a finalização.

A aplicação foi projetada com foco em organização, rastreabilidade e redução de controles manuais.

## Funcionalidades

- Autenticação de usuários
- Controle de acesso por função
- Cadastro de novas entregas
- Registro de recebimentos
- Cadastro e consulta de peças
- Preenchimento automático de descrições através da base de referências
- Conferência de quantidades
- Identificação de divergências
- Registro de avarias
- Controle de adequação da embalagem
- Anexação de NF-e
- Conferência da NF-e
- Aprovação ou rejeição da NF-e
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
     +---- Rejeitada
     |
     v
Conferência Aprovada
     |
     v
Entrada Realizada
     |
     v
Finalização
