# 📊 BH Digital - Extração de Dados para Excel  

Este projeto foi desenvolvido para **automatizar a extração de dados de uma base Oracle e salvá-los em arquivos Excel** de forma **rápida e eficiente**.  

A cada mês, consultas SQL são executadas e os resultados são organizados em arquivos mensais, cobrindo vários anos de dados. Esse processo elimina a necessidade de extração manual, tornando a tarefa mais prática e confiável. 🚀  

## 🚀 Funcionalidades  
✔️ Execução automática de consultas SQL mensalmente  
✔️ Geração de arquivos Excel organizados por mês e ano  
✔️ Facilidade na análise e organização dos dados  
✔️ Elimina a necessidade de extração manual  

## 🛠️ Tecnologias Utilizadas  

### 📌 **Linguagem & Ferramentas**  
- **TypeScript**   
- **Node.js** 

### 📌 **Banco de Dados**  
- **oracledb** - Conector para bancos Oracle  

### 📌 **Manipulação de Arquivos**  
- **exceljs** - Criação e manipulação de arquivos Excel (.xlsx)  

## 📦 Como Usar  

### 1️⃣ Clonar o repositório  
```sh
git clone https://github.com/seu-usuario/bh-digital.git
cd bh-digital
```

### 2️⃣ Instalar as dependências

```sh
npm install
```
### 3️⃣ Configurar a conexão com o banco

Edite o arquivo .env (se necessário) e configure as credenciais do banco Oracle.

### 4️⃣ Executar o script

```sh
npm run dev
```


## 📄 Estrutura do Projeto

```bash
bh-digital/
│── src/
│   ├── generate-xlsx.ts  # Script principal para geração dos arquivos Excel
│── package.json          # Configuração do projeto e dependências
│── tsconfig.json         # Configuração do TypeScript
│── .eslintrc.json        # Regras de linting
│── .gitignore            # Arquivos ignorados pelo Git
```

