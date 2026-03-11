Salvar/Carregar dados localmente (pasta)

Objetivo
- Permitir que a aplicação grave e leia o arquivo JSON `estoqueArmasV2.json` em `C:\controle estoque`.

Como funciona
- Um pequeno servidor Node.js (Express) expõe endpoints:
  - `POST http://localhost:3000/save`  -> grava JSON recebido em `C:\controle estoque\estoqueArmasV2.json`
  - `GET  http://localhost:3000/load`  -> retorna o conteúdo do arquivo salvo

Requisitos
- Node.js (v14+)
- Permissão para criar a pasta `C:\controle estoque` (usuário local normalmente possui)
- Porta `3000` livre

Instalação e execução
1. Abra o PowerShell como administrador (recomendado para criar a pasta na raiz C: sem problemas)
2. Navegue até a pasta do projeto:

```powershell
cd "e:\MEUS DOCUMENTOS\OneDrive\Documentos\Sistemas\Estoque\Controle-Estoque"
```

3. Instale as dependências (uma vez):

```powershell
npm init -y
npm install express body-parser
```

4. Execute o servidor:

```powershell
node save_server.js
```

5. No navegador, abra a aplicação `index.html` normalmente (duplo-clique) e use os botões:
   - `Salvar na Pasta` para enviar os dados atuais ao servidor e gravar em `C:\controle estoque\estoqueArmasV2.json`
   - `Carregar da Pasta` para ler o arquivo salvo e substituir os dados em memória/localStorage

Notas de segurança
- O servidor criado é um utilitário local para uso em rede local/PC. Não exponha `save_server.js` diretamente à Internet.
- Se desejar, é possível proteger com autenticação básica e regras de firewall.

Problemas comuns
- Porta 3000 em uso: altere `PORT` em `save_server.js` e em `app.js` (funções `saveToFolder`/`loadFromFolder`).
- Erro de permissão ao criar a pasta: execute PowerShell como administrador ou crie manualmente `C:\controle estoque`.

*** FIM ***
