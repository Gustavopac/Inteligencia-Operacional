# Arrow | Central de Inteligência Operacional

Painel estático para GitHub Pages: HTML, CSS e JavaScript. Não exige servidor próprio nem Power BI.

## Funcionalidades
- Visão executiva com 8 KPIs e gráficos de evolução, status, clientes, materiais, operações e tipologia.
- Operações com filtros e tabela detalhada.
- Tempos e variações: média, mediana, desvio padrão e coeficiente de variação.
- Fora da média: identificação estatística por operação + material.
- Programação por data com impressão em A4 paisagem.
- Qualidade dos dados: completude por coluna.
- Importação de Excel local, exportação CSV filtrado, login Microsoft e atualização automática a cada hora.
- Demonstração inicial com **dados fictícios**, claramente identificados.

## Publicar no GitHub Pages
1. Envie `index.html`, `style.css`, `script.js` e a pasta `Assets` para a raiz do repositório.
2. Em Settings → Pages selecione Deploy from a branch, `main`, `/ (root)`.
3. Acesse a URL do GitHub Pages. A demonstração funciona imediatamente.
4. Para dados reais, clique em **Importar Excel** ou configure o OneDrive.

## Configurar OneDrive
No início do `script.js`, altere `CFG.clientId`, `CFG.tenant`, `CFG.filePath` e `CFG.sheet`.
No Microsoft Entra ID registre um aplicativo SPA, configure a URL completa do GitHub Pages como Redirect URI e conceda a permissão delegada Microsoft Graph `Files.Read` (com consentimento aplicável). Cadastre também `http://localhost:5500/` para testes locais, se necessário.
O arquivo deve estar no OneDrive da conta autenticada, no caminho configurado. O endpoint usa a planilha do Excel via Microsoft Graph.

## Colunas
`DATA`, `CLIENTE`, `OPERAÇÃO`, `MATERIAL`, `TIPOLOGIA`, `STATUS`, `AGENDADO?`, `TEMPO`, `DESTINO`, `NF`, `OBS`; opcionalmente `ÍNICIO` e `TÉRMINNO` para calcular duração quando `TEMPO` não existe. Os cabeçalhos são normalizados sem acentos. Valores de TEMPO numéricos são interpretados como minutos (frações de dia do Excel são convertidas em minutos).

## Atenção
- Dados de demonstração **não são reais**.
- O painel usa Chart.js, SheetJS e MSAL via CDN, portanto requer internet.
- O site estático não compartilha dados de login com visitantes: cada usuário precisa autenticar e ter permissão de leitura no arquivo.
- Não coloque tokens, senhas, client secrets ou planilhas com informações sigilosas no repositório.
- GitHub Pages não é mecanismo de controle de acesso; consulte a política de segurança da empresa antes de publicar informações operacionais.
- A integração com o Microsoft Graph precisa ser validada com a planilha e as permissões reais.


## Identidade visual

Logo oficial com fundo transparente, ampliada na navegação. Gráficos e destaques usam exclusivamente gradações de laranja com cor-base `#F46E00`; fundos em preto/grafite, sem tons marrons.
