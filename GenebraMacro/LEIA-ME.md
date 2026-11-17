# Macro Gerador de Mecanismo Genebra (SOLIDWORKS 2024)

**Autor:** Gilmar Barbosa · **Autom**

O formulário mostra o logo no rodapé.

- **Gravado no `.swp` (recomendado):** no editor VBA, abra o desenho do `frmGenebra` (Shift+F7) e insira um controle **Image** pela Caixa de ferramentas. Nas Propriedades, ajuste: **(Name)** = `imgLogo`; **Picture** = `Autom_Logo.bmp` (botão "…"); **PictureSizeMode** = `3 - fmPictureSizeModeZoom`; **BorderStyle** = `0 - fmBorderStyleNone`. Depois salve. A imagem fica dentro do `.swp`, e o código ajusta posição e tamanho sozinho.
- **Sem a imagem de design:** a macro carrega `Autom_Logo.bmp` da pasta do `.swp` ou, se não encontrar, de `Imagens`. As peças e a montagem geradas recebem as propriedades personalizadas **Autor** e **Empresa**.

## Arquivos

| Arquivo | O que é |
|---|---|
| `modGenebra.bas` | Módulo principal: cálculo, validação, peças, montagem, mates, animação, CSV |
| `clsEvento.cls` | Classe que liga os eventos dos controles criados por código |
| `frmGenebra_codigo.txt` | Código do formulário, para colar num UserForm vazio |

Os três arquivos contêm **só caracteres ASCII**. Os textos acentuados estão escritos como `U("Saída")` e a função `U()` do módulo gera os acentos quando a macro roda. Por isso a codificação do arquivo e a forma de copiar não importam mais.

- `modGenebra.bas` e `clsEvento.cls` devem ser **importados** (Arquivo > Importar arquivo), não colados: as linhas `Attribute` e `VERSION` do topo só valem na importação.
- `frmGenebra_codigo.txt` é para **copiar e colar** no código do UserForm.

## Instalação (uma vez)

1. No SOLIDWORKS: **Ferramentas > Macro > Nova…** e salve como `Genebra.swp`. O editor VBA abre.
2. No editor VBA: **Inserir > UserForm**. Na janela Propriedades (F4), mude **(Name)** para `frmGenebra`.
   - Esse passo também adiciona a referência *Microsoft Forms 2.0*, necessária para a classe.
3. **Arquivo > Importar arquivo…** e importe `clsEvento.cls`. Depois importe `modGenebra.bas`.
4. Dê um duplo clique no `frmGenebra`, aperte **F7** (código), apague tudo e cole o conteúdo inteiro de `frmGenebra_codigo.txt`.
5. Apague o módulo criado automaticamente (`Module1` ou `Genebra1`), porque ele tem outro `main`.
6. Em **Ferramentas > Referências**, confirme que estão marcadas:
   - *SldWorks 2024 Type Library*
   - *SOLIDWORKS 2024 Constant type library*
   - *Microsoft Forms 2.0 Object Library*
7. **Depurar > Compilar VBAProject** para verificar se não há erros. Depois salve.
8. Para usar: **Ferramentas > Macro > Executar > Genebra.swp** e escolha `main`.

Os modelos padrão de peça e montagem precisam estar definidos em *Ferramentas > Opções > Opções do sistema > Modelos padrão*. Outra opção é informar o caminho dos templates na aba **Saída**.

## Uso

1. Preencha as abas **Geral**, **Roda**, **Manivela** e **Base e eixos**. A aba **Resultados** recalcula a cada alteração e mostra a geometria, os tempos, as velocidades e acelerações, e os erros e avisos.
2. Na aba **Saída**, escolha a pasta e clique em **GERAR**.
3. A macro cria as peças, a montagem, `Cinematica.csv` e `LEIA-ME_Motion.txt`.
4. Com a montagem aberta e ativa, use a aba **Movimento**:
   - o controle deslizante posiciona a manivela;
   - o botão **Animar** gira roda e manivela em sincronismo pela cinemática exata.

## Como a montagem funciona

- A **base é fixa**. Ela contém os eixos `EIXO_R` e `EIXO_M` e os planos `PL_RODA` e `PL_MAN`.
- Roda e manivela têm um **mate coincidente de eixo** e um **mate coincidente de plano** cada. Assim, só a rotação fica livre.
- Os eixos (peças) são acoplados aos mesmos eixos e à face da base.
- O movimento intermitente do Genebra não pode ser descrito com mates comuns, porque o pino entra e sai da ranhura. O sincronismo vem de duas fontes:
  1. **Animação da macro:** cinemática exata, funciona em qualquer edição do SOLIDWORKS.
  2. **Estudo de movimento** (Motion): motor rotativo na manivela e contato entre roda e manivela. O passo a passo, com os valores do seu projeto, fica em `LEIA-ME_Motion.txt`.

## Simplificações desta versão

- Só o Genebra **externo**, com pino integrado à manivela. Não há pino com rolamento.
- **Não há** filetes ou chanfros nas bordas das ranhuras. Aplique à mão, se quiser.
- A **Espessura da ponta do garfo** (aba Roda) define a parede mais fina entre a ranhura e o arco de travamento, que fica na própria ponta. Com Rs automático, `Rs = r − w/2 − espessura − folga do arco`. Cada mm a mais na ponta tira 1 mm de Rs. Com Rs manual, a aba Resultados mostra a espessura real e avisa se ficar abaixo da pedida. Presets antigos com "Margem arco/ranhura": a espessura equivalente é margem − folga do arco.
- A profundidade da ranhura é sempre a máxima segura: o centro do pino passa a `C − r` do centro da roda, mais a folga no fundo.
- Material: o nome precisa ser **exatamente** o da biblioteca do SOLIDWORKS em português. Se não encontrar, a macro só registra um aviso.

## Se algo falhar

O código **não foi testado dentro do SOLIDWORKS**: foi escrito sem acesso ao programa. Pontos que podem precisar de ajuste:

| Sintoma | O que verificar |
|---|---|
| "Template padrão não definido" | Informe os caminhos `.prtdot` e `.asmdot` na aba Saída |
| Falha em extrusão ou corte | Mensagem na aba Resultados; confira os avisos de geometria |
| Peças viradas/deslocadas na montagem | Aviso "não foi possível confirmar o lado do plano" no log |
| A animação não mexe as peças | Marque "Reconstruir a cada passo" |
| A roda gira para o lado errado na animação (pino bate) | Troque o sinal em `DetectarSinalRotacao` (`G_RotSign`) |
| O SOLIDWORKS fecha (crash) ao gerar | Veja a última linha de `%TEMP%\Genebra_log.txt`: ela mostra a etapa onde parou |

**Formulário não modal:** o modelo pode ser girado e ampliado com o formulário aberto. Deixe o SOLIDWORKS ao lado para acompanhar a geração. Na animação, **Parar** ou **ESC** interrompe.

**Crash resolvido (`c0000028`):** era causado por chamadas *late-bound* à MathUtility com arrays. Agora as variáveis são tipadas, e com ângulo inicial 0° a rotação nem é chamada. Se um dia o SOLIDWORKS cair, a última linha de `%TEMP%\Genebra_log.txt` mostra a etapa.

**Importante:** use sempre o código de `frmGenebra_codigo.txt`. Um `frmGenebra.frm` exportado de uma versão antiga pode ter os acentos já trocados por `?`.

Se der erro, me mande a mensagem e a linha que o VBA destacar.
