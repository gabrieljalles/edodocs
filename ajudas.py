"""Textos dos botões ⓘ do assistente de processo: o que é, como fazer e exemplos, em
linguagem simples, para quem nunca montou um modelo. Cada ajuda: (título, texto em Markdown)."""

AJUDAS = {
    "comecar": ("Como funciona um processo", """
Um **processo** é um tipo de ato que você faz muitas vezes (por exemplo, *Licença Prêmio* ou
*Compensação*). Você ensina o Edodocs **uma vez** e depois ele monta e envia os atos sozinho.

Para ensinar, você só precisa de **duas coisas**:

1. **A planilha** com os dados das pessoas: uma linha para cada ato.
2. **Um ato que você já fez** no Word, de uma das pessoas da planilha.

O Edodocs compara o ato com a linha da pessoa na planilha, descobre onde entram os dados
(nome, matrícula, dias...) e cria o modelo sozinho. Você só confere.

> Leva uns **5 minutos**. Na dúvida, clique em **Treinar com um exemplo**: nada é salvo.
"""),
    "nome": ("Nome do processo", """
É só um nome para **você reconhecer** o processo na lista. Ele aparece no canto de cima do Edodocs.

**Bons exemplos:** `Licença Prêmio` · `Compensação` · `Exoneração 2026`

Não use estes caracteres: `/ \\ : * ? " < > |`
"""),
    "planilha": ("A planilha", """
**Não tem planilha?** Clique em **Criar a planilha para mim**: o Edodocs cria uma com uma
coluna para cada campo do seu Word. As colunas de opções (ex.: *A PEDIDO*) já vêm com uma
**lista para escolher** (SIM/NÃO) e as datas, com formato de data. Você só preenche.

**Já tem?** Escolha o arquivo. Ela precisa ter os dados de quem vai receber o ato:

- **Uma linha para cada ato** (cada pessoa).
- **A 1ª linha tem os títulos das colunas** (NOME, MATRÍCULA, CARGO...).
- Pode ter outras colunas que o ato não usa: elas são ignoradas.

**Exemplo:**

| Nº DO ATO | NOME | MATRÍCULA | CARGO | DIAS | SEXO |
|---|---|---|---|---|---|
| 101/2026 | MARIA DA SILVA | 12.345-6 | AUXILIAR ADMINISTRATIVO | 15 | F |
| 102/2026 | JOÃO SOUZA | 23.456-7 | MOTORISTA | 30 | M |

**Dicas**

- A planilha **continua onde está** (no computador ou numa pasta da rede). O Edodocs lê sempre
  a versão mais nova: é só continuar preenchendo nela.
- Uma coluna **SEXO** com F/M permite trocar sozinho "lotada/lotado", "à servidora/ao servidor".
- Evite **linhas em branco no meio** e **células mescladas** no cabeçalho.
"""),
    "ajustes": ("Folha e coluna do nome", """
**Qual folha da planilha?** — a aba lá embaixo no Excel (ex.: *Plan1*, *atos*). Troque se os
dados estiverem em outra aba.

**Qual coluna tem o nome da pessoa?** — o Edodocs usa essa coluna para saber **quais linhas são
atos**: linha com nome = um ato; linha sem nome = ignorada (ex.: linhas de total ou em branco).

Na dúvida, deixe como está: o Edodocs já escolhe a coluna chamada **NOME**, se houver.
"""),
    "ato": ("Que Word devo escolher?", """
O Word (**.docx**) de **um ato que você já fez**, do jeito que foi publicado.

**Marcando você mesmo** (não precisa de planilha): o texto do Word aparece na tela.
1. Selecione com o mouse o que muda de um ato para outro (ex.: **ANTONIO EXEMPLO DOS SANTOS**).
2. Clique em **Marcar o que selecionei** e dê um nome: **NOME**.
3. Repita para matrícula, cargo, data...
4. O que tem opções (ex.: *"a pedido, "* que às vezes aparece e às vezes não) marque como
   **Escolha entre opções**: SIM escreve *"a pedido, "*, NÃO não escreve nada.

O Edodocs também sugere trechos que parecem dados (nomes em maiúsculas, matrículas, datas):
é só clicar na sugestão.

**Com a planilha:** clique em **Usar a minha planilha** e escolha o ato de uma pessoa que está
nela. O Edodocs procura no texto os valores da linha dela ("MARIA DA SILVA", "12.345-6"...) e
marca sozinho.

**Funciona melhor quando:**

- os dados no Word estão **escritos igual** à planilha (mesmo nome, mesma matrícula);
- o título do ato está na **primeira linha** do Word.

Já tem um modelo com campos `{{NOME}}`? Pode escolher ele aqui também: o Edodocs reconhece.
"""),
    "achados": ("Marcar ou desmarcar?", """
Cada linha mostra um trecho do ato com o dado encontrado **destacado**.

- **Marcado** ✅ = muda de uma pessoa para outra → vira um campo do modelo.
- **Desmarcado** ☐ = é texto fixo do ato → fica sempre igual.

**Exemplos:**

| Trecho | O que fazer |
|---|---|
| ...à servidora **MARIA DA SILVA**, matrícula... | ✅ marcado (é o nome de cada pessoa) |
| ...conceder **15** (quinze) dias... | ✅ marcado (os dias mudam) |
| ...com base na Lei nº **340**... | ☐ desmarque (a lei é sempre a mesma) |

**"É a coluna"**: de qual coluna da planilha vem o dado. Se o mesmo valor aparece em duas
colunas (ex.: 15 em DIAS e em PADRÃO), confira e troque se precisar.
"""),
    "modelo_manual": ("Preparar o modelo você mesmo", """
Se preferir, prepare o Word do modelo:

1. Escreva o ato normalmente, com a formatação que deve sair.
2. Onde entra um dado da planilha, escreva o **nome da coluna entre chaves duplas**.
3. Deixe o título na primeira linha.

**Exemplo:**

> a) Conceder {{DIAS}} dia(s) de licença à servidora {{NOME}}, matrícula nº {{MATRÍCULA}}...

Campos de mala direta do Word («NOME») também funcionam.
"""),
    "campos": ("De onde vem cada campo", """
Cada campo do modelo precisa saber **de onde vem o valor**. A bolinha 🟢 diz que está certo;
🔴 diz que falta escolher algo. O **Ex.:** mostra o valor de uma pessoa da planilha.

| Tipo | Para que serve | Exemplo |
|---|---|---|
| **Coluna da planilha** | copia o valor da coluna | NOME → *MARIA DA SILVA* |
| **Texto fixo** | sempre o mesmo texto | *Secretaria Municipal de Administração* |
| **Número por extenso** | escreve o número em palavras | DIAS = 15 → *quinze* |
| **Feminino / masculino** | troca a palavra pelo sexo (coluna com F/M) | F → *lotada* · M → *lotado* |
| **Data por extenso** | escreve a data em palavras | 05/10/2026 → *5 de outubro de 2026* |
| **Muda conforme a planilha** | regras "Se ... escrever ..." | A PEDIDO é SIM → *a pedido,* |

**Se vazio, usar**: o texto que entra quando a célula da planilha estiver vazia (opcional).

**Muda conforme a planilha**: clique em **Editar regras** e monte, escolhendo nas listas:

> **Regra 1:** Se *A PEDIDO* **é** *SIM* e *VÍNCULO* **é** *EFETIVO* → escrever *artigo 48*
> **Regra 2:** Se *A PEDIDO* **é** *SIM* → escrever *artigo 49, II*
> **Se nenhuma servir:** escrever *inciso I do artigo 49*

O Edodocs usa a **primeira** regra que servir. O texto pode ter dados da planilha
(*retroativo a {{DATA}}*) e pode ficar **vazio**: se um parágrafo do Word só tinha esse campo
(ex.: *{{CONSIDERANDO}}*), ele some do ato.
"""),
    "varias": ("Mais de uma pessoa no mesmo ato", """
Alguns atos têm **várias pessoas** (ex.: *exonera FULANO e BELTRANO*), cada uma num parágrafo:

> I – PAULO..., da Secretaria de Governo;
> II – LUIZ..., da Secretaria de Desenvolvimento Social.

No passo do Word, clique em **Escolher o parágrafo de uma pessoa** e escolha o do PAULO. O
Edodocs repete esse parágrafo para cada pessoa, troca o "I" pela numeração automática e tira
o parágrafo do LUIZ (ele volta sozinho, com os dados da planilha).

Por trás, funciona assim:

1. **Na planilha**, cada pessoa continua numa linha, e uma coluna diz quais vão juntas
   (ex.: *Nº DO DECRETO* ou *GRUPO*): **linhas com o mesmo valor viram um ato só**.
   Linha com essa coluna vazia vira um ato sozinha.
2. **No Word**, clique em **Escolher o trecho** e marque o parágrafo de cada pessoa
   (o que tem *{{NOME}}*, *{{MATRÍCULA}}*...). O Edodocs repete esse trecho para cada uma.
3. Nesse trecho, use os campos que o Edodocs preenche sozinho:

| Campo | Vira |
|---|---|
| {{ITEM}} | I, II, III... |
| {{Nº DO ITEM}} | 1, 2, 3... |
| {{LETRA DO ITEM}} | a, b, c... |
| {{FIM DO ITEM}} | **;** nas pessoas do meio e **.** na última |
| {{QUANTIDADE DE PESSOAS}} | 1, 2, 3... (vale no ato todo) |

**Exemplo** de trecho no Word:

> {{ITEM}} – {{NOME}}, matrícula nº {{MATRÍCULA}}, do cargo de {{CARGO}}{{FIM DO ITEM}}

vira

> I – MARIA DA SILVA, matrícula nº 12.345-6, do cargo de Assistente;
> II – JOÃO SOUZA, matrícula nº 23.456-7, do cargo de Coordenador.

**Fora do trecho**, os campos usam a **primeira pessoa** do ato. Para mudar o texto quando
todos são da mesma secretaria, use um campo **Muda conforme a planilha** com a condição
*SECRETARIA* **é igual para todas as pessoas do ato**.
"""),
    "colunas": ("Colunas da tabela dos atos", """
Aqui você escolhe **como a tabela da página Atos vai aparecer**:

- **Setas** ↑ ↓: mudam a ordem das colunas.
- **Olho**: mostra ou esconde a coluna (ela continua na planilha).
- **Nova caixinha**: cria uma coluna de **marcar à mão**, com o nome que você quiser
  (ex.: *CONFERIDO*, *ASSINADO*, *ARQUIVADO*). Na tabela dos atos, cada linha ganha uma
  caixinha; ao marcar, o Edodocs grava **SIM** nessa coluna da planilha.

Dá para mudar tudo isso depois, no botão **Colunas** da página Atos.
"""),
    "outros_modelos": ("Um Word para cada situação", """
Às vezes o ato de **uma pessoa** e o de **várias** são escritos de jeitos bem diferentes:

> *Art. 1º Fica exonerado FULANO, matrícula...*
> *Art. 1º Ficam exonerados os servidores abaixo, sendo: I – ...; II – ....*

Em vez de complicar um Word só, **faça um Word para cada jeito** e diga quando usar cada um:

- O **Word principal** (o do passo 2) é o padrão.
- **Adicionar outro Word...** e, em **Quando usar**, escolha a condição.
  Ex.: *QUANTIDADE DE PESSOAS* **é maior que** *1*.

O Edodocs escolhe o Word de cada ato sozinho. Os campos de todos os Words aparecem aqui
embaixo para você ligar.
"""),
    "titulos": ("Títulos do ato", """
O **título** identifica o ato nos sites:

- **No e-docs**: vai no campo *Nossa identificação*.
- **No Diário Oficial**: é o título da publicação.

O Edodocs tira o título da **primeira linha do Word** e já coloca o número de cada ato.
Confira acima, com a pessoa escolhida, se está certo.

**Para mudar** (em *Ajustar títulos*):

- `{NOME DA COLUNA}` vira o valor da coluna. Ex.: `ATO SMA Nº {Nº DO ATO}`.
- `{número do ato}` = o número da coluna do número (ex.: *1542/2026*).
- `{ano do ato}` = o ano do número do ato (ex.: *2026*).
"""),
    "pessoas": ("Vistos e assinantes", """
**Sempre os mesmos vistos e assinantes**: ligue se, neste processo, quase sempre vistam e assinam
as mesmas pessoas. Assim o Edodocs não pergunta a cada envio.

- Quem envia **só vista se estiver na lista**: para vistar os próprios atos, inclua o seu nome.
  Você vista sempre **em 1º**, na frente dos outros.
- Escolha as pessoas pesquisando por **CPF, nome ou matrícula** (com o dígito no final).

Deixe desligado se muda muito: ele pergunta na hora de enviar. Dá para mudar depois em
**Processos → Vistos e assinantes**.
"""),
    "envio": ("Perguntar antes de enviar", """
**Ligado** (recomendado no começo): antes de cada envio o Edodocs mostra o ato pronto e você
escolhe **Enviar**, **Pular este ato** ou **Cancelar tudo**. Bom para acompanhar e conferir.

**Desligado**: os atos são enviados **um atrás do outro, sem parar**. Bom para processos que você
já confia e tem muitos atos. Você acompanha tudo pela página **Execução** e pode cancelar lá.

Cada processo tem a sua escolha. Dá para mudar a qualquer momento em **Processos**.
"""),
    "diario": ("Publicado no Diário Oficial?", """
**Sim** (o comum): depois de enviado no e-docs, o ato é **cadastrado no Diário Oficial** e o
Edodocs **confere a publicação**.

- O Diário recebe o ato **sem a primeira linha do Word**, porque ela é o título e vai num
  campo próprio do site (veja a aba *No Diário Oficial* ao lado).

**Não**: o processo só vai ao **e-docs**. As etapas *Lançar no Diário* e *Conferir publicação*
ficam desligadas para ele.

Marcou errado? Dá para mudar a qualquer momento em **Processos** (no cartão do processo).
"""),
    "sites": ("Opções dos sites", """
Como o ato é cadastrado nos sites. **Só mude se o seu ato for diferente** do padrão.

| Opção | O que é | Padrão |
|---|---|---|
| e-docs: Setor | um dos setores que aparecem **para você** no e-docs | o 2º da sua lista |
| e-docs: Tipo | um dos tipos que aparecem para você naquele setor | ATOS ADMINISTRATIVOS |
| Diário: secretaria | quem publica | Administração |
| Diário: categoria | categoria da publicação | Ato Administrativo |
| Diário: coluna | o layout da publicação | Duas colunas |

No e-docs, **Setor** e **Tipo** são listas com as opções do **seu** usuário: o Edodocs busca no
site (cada pessoa tem as suas). Se mudar o setor, a lista de tipos é buscada de novo.

Na dúvida, veja no site como você faz à mão e escolha o mesmo aqui.
"""),
}
