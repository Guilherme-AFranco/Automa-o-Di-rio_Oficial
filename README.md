# Automação - Diário Oficial

**Instalação do ChromeDriver**
- Verifique a versão do seu navegador Google Chrome em: ... - Ajuda - Sobre o Google Chrome;
- Acesse [Chromedriver]([https://chromedriver.storage.googleapis.com/index.html?path=114.0.5735.90/](https://sites.google.com/chromium.org/driver/));
- Faça download do .zip;
- Extraia em um local conhecido;
- Dentro do Windows Explorer, clique com o botão direito em "Meu Computador" e selecione Propriedades;
- Vá em Configurações Avançadas do Sistema > Variáveis de Ambiente;
- Na seção Variáveis do Sistema, localize Path, clique em Editar;
- Adicione o caminho completo da pasta do ChromeDriver, por exemplo: C:\WebDriver;
- Salve.

**Criação e uso do ambiente virtual**
- Abra o VSCode e entre na pasta onde estão os arquivos do Dashboard;
- No canto superior direito da tela haverá uma opção "Toggle Panel (Ctrl + J)". Abra-o;
- Dentro do terminal que será aberto, digite "python -m venv venv" para criação do seu ambiente virtual na pasta em que estiver selecionada (faça isso apenas no primeiro acesso);
- Acesse-o utilizando o comando "\venv\Scripts\activate" e digite r para permitir (faça isso sempre que abrir o VSCode);
- Instale o arquivo `requirements.txt` dentro do seu ambiente virtual a partir do comando "pip install -r .\requirements.txt" (faça isso apenas no primeiro acesso).

**Criação da chave API do GEMINI**
- Acessar o site [Google Developer AI](https://ai.google.dev/);
- Entre em "Get API key in Google AI Studio";
- Crie a chave de API e copie;
- Incula a chave em um documento `.env` da mesma forma que o modelo `.env_example`.
