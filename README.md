# RITZ · Painel

Painel operacional e financeiro da loja RITZ. Ele lê ao vivo a planilha **controle-gerencial** (Google Drive) através de um Web App do Apps Script.

- `index.html`: o painel. Não tem dados nem credenciais. O link `/exec` e a senha ficam salvos no navegador de cada pessoa.
- `apps-script/Code.gs`: a ponte entre a planilha e o painel (é colado na planilha em Extensões → Apps Script).

Sem o link e a senha, a página não mostra nenhum número.
