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
- Consulta de referências através da base de dados
- Preenchimento automático da descrição das peças
- Conferência de quantidades
- Identificação de divergências
- Registro de avarias
- Controle da adequação da embalagem
- Anexação de NF-e
- Visualização da NF-e no sistema
- Conferência da NF-e
- Aprovação ou rejeição da NF-e
- Controle de entradas pendentes
- Histórico das movimentações
- Gerenciamento de usuários
- Controle de cargos e permissões
- Alteração e redefinição de senhas
- Integração com banco de dados SQLite
- Integração com arquivo Excel para consulta de referências

## Fluxo do processo

O processo operacional segue as seguintes etapas:

**Nova Entrega**  
↓  
**Recebimento**  
↓  
**Conferência Física**  
↓  
**Anexação da NF-e**  
↓  
**Conferência da NF-e**  
↓  
**Aprovação ou Rejeição**

- **Aprovado** → Entrada Realizada → Finalização
- **Rejeitado** → Retorno para correção

---

## Interface do sistema

### Tela de Login

![Tela de Login](./capturas%20de%20tela/01-login.png)

### Painel Principal

![Painel Principal](./capturas%20de%20tela/02-painel.png)

### Nova Entrega

![Nova Entrega](./capturas%20de%20tela/03-nova-entrega.png)

### Receber Entrega

![Receber Entrega](./capturas%20de%20tela/04-receber-entrega.png)

### Entradas Pendentes

![Entradas Pendentes](./capturas%20de%20tela/05-entradas-pendentes.png)

### Conferência da NF-e

![Conferência da NF-e](./capturas%20de%20tela/06-conferencia-nfe.png)

### Entrada do Recebimento

![Entrada do Recebimento](./capturas%20de%20tela/07-entrada-recebimento.png)

### Notas Fiscais

![Notas Fiscais](./capturas%20de%20tela/08-notas-fiscais.png)

### Gerenciamento

![Gerenciamento](./capturas%20de%20tela/09-gerenciamento.png)

## Linguagens e tecnologias utilizadas

### Linguagens

- **Python** — desenvolvimento do backend e regras de negócio
- **HTML5** — estrutura das páginas e interfaces
- **CSS3** — estilização e layout do sistema
- **JavaScript** — interações e funcionalidades no frontend

### Tecnologias e ferramentas

- **Flask** — framework utilizado para desenvolvimento da aplicação web
- **SQLite** — banco de dados do sistema
- **OpenPyXL** — integração e leitura da base de dados em Excel
- **JSON** — armazenamento e controle de informações do fluxo
- **Git e GitHub** — versionamento e gerenciamento do projeto
