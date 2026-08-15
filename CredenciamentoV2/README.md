# Sistema de Gestão de Evento (100% gratuito — Python/Flask)

Painel de inscrição + credenciamento + check-in/check-out com horário exato,
rodando localmente no seu computador. Sem custos, sem serviços externos,
sem limite de participantes.

## O que ele faz

- **Tela inicial** (`/`): logo do evento, cronograma completo dos dois dias
  e botão "Fazer Inscrição" que abre o formulário em outra aba.
- **Inscrição pública** (`/inscricao`): nome, tipo (aluno ou participante
  externo), e-mail (com aviso pedindo e-mail institucional quando for aluno),
  formação/curso e instituição (faculdade, empresa etc.).
- **Credencial automática**: ao se inscrever, o participante recebe um QR Code
  único e um número de inscrição sequencial (202601, 202602, 202603...). A
  credencial pode ser baixada em PDF, com QR Code + nome + número de
  inscrição, pronta para impressão ou para guardar no celular.
- **Check-in / Check-out** (`/scanner`, protegido por senha): a câmera só é
  ativada quando a equipe toca em "Ler QR Code" (não liga sozinha). O sistema
  alterna automaticamente entre ENTRADA e SAÍDA e grava a data/hora exata do
  servidor. Se o QR Code não conseguir ser lido, dá para liberar o acesso
  digitando o número de inscrição da pessoa.
- **Painel** (`/painel`, protegido por senha): total de inscritos, quantos
  estão presentes agora, quantos são alunos x participantes externos, e o
  histórico de cada participante (incluindo tipo e instituição).
- **Exportação CSV** (`/exportar`, protegido por senha): abre direto no
  Excel/Google Sheets, com uma linha por evento de entrada/saída e o horário
  de cada um — ideal para calcular horas de permanência para certificados.

## Dois níveis de acesso: ADM e equipe normal

A inscrição (`/inscricao`) continua **totalmente aberta**, sem senha —
qualquer aluno se inscreve livremente. Já a área da equipe (`/login`) agora
tem **duas senhas diferentes**, dando acessos diferentes:

| Acesso | Senha padrão | Pode acessar |
|---|---|---|
| **ADM** | `ADMUNiTALKS2026` | Check-in/check-out (`/scanner`) **+** Painel (`/painel`) **+** Exportar CSV (`/exportar`) |
| **Equipe normal** | `unitalks2026` | Somente check-in/check-out (`/scanner`) |

Quem loga com a senha normal **não vê o link do Painel no menu** e, se
tentar acessar `/painel` ou `/exportar` diretamente pela URL, é redirecionado
de volta ao scanner com um aviso — ou seja, é bloqueado mesmo digitando o
endereço na mão.

- **Para trocar as senhas**: defina as variáveis de ambiente
  `ADMIN_PASSWORD` e `ACCESS_PASSWORD`
  - No Render: vá em seu serviço → aba **Environment** → **Add Environment
    Variable** → adicione as duas chaves com os valores que você quiser.
  - Localmente (Windows, no `cmd`, antes de rodar):
    ```
    set ADMIN_PASSWORD=suasenhaadm
    set ACCESS_PASSWORD=suasenhaequipe
    python app.py
    ```

**Como funciona na prática**: cada celular da equipe faz login uma vez
(fica conectado por ~18 horas, cobrindo o dia inteiro do evento sem pedir
senha de novo). Como cada celular guarda sua própria sessão, várias pessoas
podem estar logadas ao mesmo tempo em aparelhos diferentes, sem se
atrapalharem — uma pessoa fazendo login não desconecta a outra.

Também recomendo definir a variável `SECRET_KEY` no Render (aba
Environment) com qualquer texto aleatório longo — ela protege a sessão de
login contra falsificação.

## Colocando online de vez, grátis (Render) — recomendado

Essa é a opção mais tranquila para o dia do evento: nada para baixar, link
fixo, HTTPS de verdade (câmera funciona sem gambiarra), acessível de
qualquer rede/celular.

### Passo 1 — Colocar o projeto no GitHub (grátis)

1. Crie uma conta grátis em `https://github.com/signup` (se ainda não tiver).
2. Clique em **"New repository"** (botão verde, no canto superior direito).
3. Dê um nome, ex.: `credenciamento-evento`. Deixe como **Public**. Não
   marque nenhuma opção extra. Clique em **Create repository**.
4. Na página do repositório recém-criado, clique em
   **"uploading an existing file"**.
5. Arraste **todos os arquivos e pastas de dentro de `evento_app`**
   (não a pasta `evento_app` inteira, o conteúdo dela) para essa área.
6. Role para baixo e clique em **"Commit changes"**.

### Passo 2 — Criar o serviço no Render

1. Crie uma conta grátis em `https://dashboard.render.com/register`
   (pode entrar direto com sua conta do GitHub).
2. Clique em **"New +"** → **"Web Service"**.
3. Selecione o repositório `credenciamento-evento` que você acabou de criar.
4. Preencha:
   - **Name**: algo como `credenciamento-evento` (vira parte do link)
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Instance Type**: **Free**
5. Clique em **"Create Web Service"**.
6. Aguarde uns 2 a 3 minutos (o Render instala tudo sozinho). Quando
   terminar, você verá um link fixo tipo:
   ```
   https://credenciamento-evento.onrender.com
   ```
7. Pronto — esse é o link definitivo. A tela de inscrição fica em
   `/inscricao`, o scanner em `/scanner`, o painel em `/painel`.

### ⚠️ Limitações do plano gratuito do Render (importante saber)

- **"Dorme" após 15 minutos sem uso**: se ninguém acessar por um tempo, o
  serviço "dorme" e a primeira pessoa a abrir o link depois disso espera
  uns 30-50 segundos até ele "acordar". Depois disso, fica rápido normal.
- **Os dados podem ser apagados se o serviço reiniciar** enquanto está
  dormindo (o banco SQLite fica no disco temporário do plano free). Veja a
  seção **"Backup automático do banco"** logo abaixo — ela existe
  justamente por causa desse risco.
- Se quiser eliminar esses limites de vez (não obrigatório), o plano pago
  mais barato do Render (a partir de uns US$7/mês) já resolve os dois
  pontos acima com um disco persistente de verdade.

### Mantendo o site sempre acordado (UptimeRobot, grátis)

1. Crie uma conta grátis em `uptimerobot.com`.
2. Clique em **"+ Add New Monitor"**.
3. Preencha:
   - **Monitor Type**: `HTTP(s)`
   - **Friendly Name**: `UniTalks 2026` (ou qualquer nome)
   - **URL (or IP)**: o link do seu site no Render, ex.:
     `https://unitalks.onrender.com`
   - **Monitoring Interval**: `5 minutes`
4. Clique em **"Create Monitor"**.

Pronto — a cada 5 minutos o UptimeRobot acessa seu site, o que impede o
Render de colocá-lo para dormir. Ative isso um pouco antes do evento
começar e pode desligar depois que acabar (sem necessidade nenhuma durante
o resto do ano, já que ele consome as horas gratuitas do Render).

### Atualizando o site depois

Se precisar alterar algo no código, é só editar o arquivo direto pelo
GitHub (ícone de lápis na página do arquivo) e salvar — o Render detecta a
mudança e atualiza o site sozinho em 1-2 minutos.

## Backup automático do banco (importante!)

O SQLite roda em disco temporário no plano gratuito do Render — ou seja,
**se o serviço reiniciar, os dados podem se perder**. Existem duas camadas
de proteção prontas no sistema:

### 1) Backup manual, sob demanda (já funciona, sem configurar nada)

No **Painel** (login de ADM), tem um botão **"💾 Baixar backup do banco"**
que baixa na hora um arquivo `.db` com tudo (inscritos + check-ins), gerado
de forma segura mesmo com o site em uso por outras pessoas ao mesmo tempo.
Recomendado: baixe esse arquivo algumas vezes durante o evento (ex.: a
cada intervalo, e no final do dia).

### 2) Backup automático por e-mail (opcional, recomendado para o dia do evento)

Se você configurar as variáveis de ambiente abaixo no Render (aba
**Environment**), o sistema envia sozinho, de tempos em tempos, um e-mail
com o backup em anexo — sem precisar lembrar de clicar em nada:

| Variável | O que colocar |
|---|---|
| `BACKUP_EMAIL_TO` | O e-mail que vai **receber** os backups (o seu, por ex.) |
| `BACKUP_SMTP_HOST` | `smtp.gmail.com` (se for usar Gmail) |
| `BACKUP_SMTP_PORT` | `587` |
| `BACKUP_SMTP_USER` | O e-mail que vai **enviar** (pode ser o mesmo do destino) |
| `BACKUP_SMTP_PASSWORD` | Uma **Senha de App** do Gmail (não é a senha normal — veja abaixo) |
| `BACKUP_INTERVAL_MINUTES` | De quanto em quanto tempo enviar (padrão: `120` = a cada 2h) |

**Como gerar a "Senha de App" do Gmail** (necessário — o Gmail bloqueia
login direto com a senha normal para esse tipo de uso):
1. Acesse `myaccount.google.com/apppasswords` (é preciso ter a verificação
   em duas etapas ativada na conta).
2. Crie uma nova senha de app, com o nome "UniTalks" por exemplo.
3. O Google mostra uma senha de 16 letras — copie e cole exatamente ela no
   campo `BACKUP_SMTP_PASSWORD` do Render (sem espaços).

Se essas variáveis não forem configuradas, o sistema simplesmente não
tenta enviar nada por e-mail — continua funcionando normalmente, com
apenas o backup manual disponível.

## Modo alternativo: rodar local + ngrok

O projeto já vem com o arquivo **`iniciar_evento.bat`**. Ele sobe o Flask e o
ngrok juntos, em duas janelas separadas, sem precisar digitar comando nenhum
no dia do evento.

**Antes do evento, configure uma vez:**
1. Instale o ngrok e autentique (veja seção abaixo).
2. Abra `iniciar_evento.bat` com o **Bloco de Notas** (clique direito →
   Editar) e ajuste a linha:
   ```
   set NGROK_PATH=C:\ngrok\ngrok.exe
   ```
   para o caminho real onde você extraiu o `ngrok.exe`.

**No dia do evento:**
- Dê duplo clique em `iniciar_evento.bat`.
- Duas janelas pretas vão abrir: uma do servidor, outra do ngrok.
- Na janela do ngrok, copie o link que aparece na linha `Forwarding`
  (algo como `https://xxxx.ngrok-free.app`) e acrescente `/scanner` no final.
- Esse é o link que vai no celular do(a) recepcionista.
- **Não feche nenhuma das duas janelas** enquanto o evento estiver rolando.

## Instalando e configurando o ngrok (fazer uma vez, antes do evento)

1. Crie uma conta grátis em `https://dashboard.ngrok.com/signup`.
2. Baixe o ngrok para Windows em
   `https://dashboard.ngrok.com/get-started/setup/windows` e extraia o
   `ngrok.exe` (sugestão: crie a pasta `C:\ngrok` e coloque ele lá).
3. Na mesma página, copie o comando de autenticação (algo como
   `ngrok config add-authtoken SEU_TOKEN`), abra o Prompt de Comando,
   navegue até a pasta do ngrok (`cd C:\ngrok`) e rode esse comando uma vez.
4. Pronto — depois disso, é só usar o `iniciar_evento.bat` sempre que for
   testar ou usar no evento.

⚠️ **No plano gratuito do ngrok**, o link muda toda vez que você reinicia o
túnel — por isso, no dia do evento, suba o `iniciar_evento.bat` **uma única
vez** e deixe as janelas abertas o dia inteiro. A primeira vez que alguém
abrir o link pelo celular, o ngrok mostra uma tela de aviso — é só clicar em
"Visit Site" uma vez.

## Como rodar (modo manual, sem o .bat)

1. Tenha Python 3.10+ instalado.
2. No terminal, dentro desta pasta:
   ```bash
   pip install -r requirements.txt
   python app.py
   ```
3. Abra `http://localhost:5000` no navegador (ou no celular, usando o IP da
   máquina na mesma rede Wi-Fi, ex.: `http://192.168.0.10:5000`).

Um arquivo `evento.db` (SQLite) é criado automaticamente na primeira execução —
é o seu banco de dados, não precisa instalar nada além do Python.

⚠️ **Se você já tinha instalado as dependências antes** (de uma versão
anterior deste projeto) e a geração do PDF der erro
`ModuleNotFoundError: No module named 'reportlab'`, é só rodar de novo:
```bash
pip install -r requirements.txt
```
Isso instala a biblioteca nova (`reportlab`, usada para gerar o PDF da
credencial) sem afetar nada do que já estava funcionando.

## Usando como "app" na recepção (sem precisar de loja de aplicativos)

A tela de credenciamento (`/scanner`) é um **PWA** (Progressive Web App):
no celular da recepção, basta abrir o link no navegador (Chrome/Safari) e
tocar em **"Adicionar à tela inicial"**. Isso cria um ícone no celular que
abre o app em tela cheia, sem barra de endereço, exatamente como um app
nativo — sem precisar publicar em loja nenhuma. Ao tocar no ícone, ele
mantém a câmera ligada continuamente, tocando um bipe e vibrando a cada
leitura, para agilizar a fila.

## Uso no dia do evento

1. Divulgue o link `http://SEU-IP:5000/inscricao` antes do evento (ou deixe um
   tablet/notebook na entrada para inscrição no local).
2. Cada participante recebe/baixa seu QR Code na tela de confirmação.
3. Na entrada, a equipe abre `http://SEU-IP:5000/scanner` no celular e escaneia
   o QR Code de cada pessoa → registra ENTRADA.
4. Na saída, escaneia de novo → registra SAÍDA automaticamente (o sistema
   detecta que a próxima leitura é uma saída).
5. A qualquer momento, acompanhe tudo em `/painel` e exporte o relatório em
   `/exportar`.

## Colocando online para acesso remoto (opcional, também gratuito)

Se quiser que as pessoas se inscrevam pela internet (não só na rede local),
você pode hospedar de graça em serviços como **Render**, **Railway** (planos
free) ou **PythonAnywhere**. Nesses casos, troque o SQLite por um caminho de
disco persistente ou, se crescer muito, migre para Postgres gratuito (Render
e Railway oferecem isso no plano free).

## Estrutura do projeto

```
evento_app/
├── app.py                # backend Flask (rotas, banco, lógica de check-in)
├── requirements.txt
├── evento.db              # criado automaticamente (SQLite)
├── static/
│   ├── css/style.css
│   └── qrcodes/            # QR Codes gerados por participante
└── templates/
    ├── base.html
    ├── inscricao.html
    ├── confirmacao.html
    ├── scanner.html
    └── painel.html
```

## Possíveis evoluções

- Autenticação simples no `/painel` (usuário/senha) antes de liberar em produção.
- Campo extra para tipo de credencial (palestrante, staff, participante).
- Emissão automática de certificado em PDF calculando horas de permanência
  (dá para usar a biblioteca `reportlab`, também gratuita).
- Múltiplos eventos no mesmo sistema (basta adicionar uma tabela `eventos` e
  vincular os participantes a ela).
