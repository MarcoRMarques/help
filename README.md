Arquivo ajuda

Autor: MRM desenvolvimentos

Programa: Sistema de Cálculo para comissões Parceiros MegaPay / Ceopag



Primeiro passo. 



Configurando corretamente o Excel para o bom funcionamento

O sistema está baseado na linguagem VBA, O VBA (Visual Basic for Applications) é uma linguagem de programação desenvolvida pela Microsoft, utilizada principalmente para automatizar tarefas em aplicativos do Microsoft Office, como o Excel, Word, Access e outros. É uma versão do Visual Basic, uma linguagem de programação mais antiga, adaptada para ser usada em ambientes de aplicativos como planilhas, documentos e bancos de dados.

Características principais do VBA:

Automatização: VBA permite criar macros, que são sequências de comandos e instruções para automatizar tarefas repetitivas. Isso economiza tempo e reduz o risco de erros ao realizar tarefas manuais.

Integração com o Microsoft Office: É amplamente utilizado dentro do Excel, Word, Access e outros produtos da Microsoft, para personalizar e estender a funcionalidade desses aplicativos.

Fácil de aprender: A sintaxe do VBA é relativamente simples, especialmente para quem já está familiarizado com o Microsoft Office. Ele permite usar funções e procedimentos para interagir com objetos e dados dos aplicativos.

Acesso a Objetos e Eventos: No VBA, você pode interagir com objetos de documentos, planilhas, gráficos e outros componentes do Office. Ele também permite responder a eventos (como cliques de botões, alterações em células de planilhas, etc.).

Portanto, ele pode afetar as interações com o ambiente Office em geral por este principal motivo, deve-se configurar o ambiente de macros para os eu perfeito funcionamento.



1º Passo – Configurando a permissão das Macros, como fazer isso?

Para garantir que as macros funcionem corretamente no Excel, você precisará configurar o Excel para permitir a execução delas. Isso envolve ajustar as configurações de segurança para permitir a execução de macros, caso elas sejam desabilitadas por padrão. Aqui está um passo a passo para configurar o Excel para usar macros:

Passo 1: Habilitar as macros no Excel

Abra o Excel.

Clique na guia "Arquivo" (ou "File" em inglês) no canto superior esquerdo.

Selecione "Opções" (ou "Options" em inglês).

Na janela "Opções do Excel", clique em "Central de Confiabilidade" (ou "Trust Center").

Clique em "Configurações da Central de Confiabilidade" (ou "Trust Center Settings").

Na janela "Central de Confiabilidade", selecione "Configurações de Macro" (ou "Macro Settings").

Você verá várias opções de configuração. Selecione uma das seguintes opções, dependendo de como você deseja configurar a execução das macros:

Desabilitar todas as macros sem notificação: As macros serão desabilitadas e o Excel não notificará você.

Desabilitar todas as macros com notificação: As macros serão desabilitadas, mas o Excel exibirá uma notificação perguntando se você deseja habilitá-las.

Habilitar macros com assinatura digital: Apenas macros de fontes confiáveis com assinaturas digitais serão habilitadas.

Habilitar todas as macros (não recomendado para segurança): Todas as macros serão executadas, independentemente de sua origem. Escolha essa opção apenas se confiar no código da macro.

Selecione a opção "Habilitar todas as macros" (se você souber que os arquivos que abrirá são de fontes confiáveis) ou "Habilitar macros com notificação" (para ser mais seguro).

Após selecionar a opção desejada, clique em "OK".

Passo 2: Habilitar a execução de macros ao abrir um arquivo

Quando você abrir um arquivo de Excel que contenha macros, o Excel pode exibir uma barra amarela no topo da planilha, dizendo "Macros desabilitadas". Para habilitar as macros nesse arquivo, clique em "Habilitar conteúdo" ou "Habilitar macros".

Passo 3: Salvar o arquivo com macros

Salve o arquivo como um arquivo habilitado para macro. Para isso, ao salvar, escolha o formato ".xlsm" (selecione "Pasta de Trabalho do Excel habilitada para macros" ao salvar o arquivo). Isso é importante, pois os arquivos no formato ".xlsx" não suportam macros.

Passo 4: Configurar a "Central de Confiabilidade" (opcional)

Se você usar macros de fontes confiáveis ou armazenar seus arquivos de Excel em locais específicos, pode configurar a Central de Confiabilidade para permitir que o Excel abra arquivos de fontes confiáveis automaticamente, sem solicitar permissão para habilitar macros.

Na janela "Central de Confiabilidade", você pode adicionar "Locais confiáveis" para que o Excel não bloqueie macros em pastas específicas. Clique em "Locais Confiáveis" e adicione a pasta onde você armazena seus arquivos com macros.













Passo 5: Certificar-se de que os macros podem ser executados em VBA

Se você estiver criando suas próprias macros em VBA, certifique-se de que o Editor VBA esteja acessível:

Para abrir o Editor VBA, pressione  Alt + F11.

Este passo é de suma importância para garantir o bom funcionamento



Ao abrir um editor do VBA, de qualquer arquivo Excel

Como acessar as referências no Editor VBA

Abra o Editor VBA:

No Excel, pressione Alt + F11 para abrir o Editor VBA.

Abrir o menu de Referências:

No Editor VBA, clique em "Ferramentas" na barra de menu.

Selecione "Referências...". Isso abrirá a janela "Referências - VBAProject".

Entendendo a janela de Referências:

A janela de Referências lista todas as bibliotecas que estão disponíveis para o seu projeto VBA. As bibliotecas selecionadas terão uma marca de seleção ao lado de seus nomes.

Você verá várias bibliotecas do sistema do Excel, como Microsoft Excel Object Library, Microsoft Forms Object Library, e outras dependendo do seu ambiente e configurações.

Adicionar ou remover referências:

Para adicionar uma nova referência, role pela lista e marque a caixa ao lado do nome da biblioteca que você deseja adicionar. Por exemplo, se você deseja usar o Microsoft Outlook Object Library para enviar e-mails, marque essa biblioteca.

Para remover uma referência, desmarque a caixa ao lado da biblioteca que você não precisa mais.

Para pesquisar uma biblioteca específica, você pode usar a caixa de pesquisa na parte superior da janela de referências.

Exemplos de bibliotecas comuns:

Microsoft Excel Object Library: Permite controlar o Excel através do VBA.

Microsoft Scripting Runtime: Oferece suporte a objetos como FileSystemObject para trabalhar com arquivos e pastas.

Este projeto necessita que estas duas referências estejam habilitadas 



Após selecionar ou desmarcar as bibliotecas desejadas, clique em "OK" para aplicar as mudanças.

Verificar a referência no código:

Uma vez que você tenha configurado as referências, o código VBA pode usar objetos e métodos dessas bibliotecas. Por exemplo, se você habilitou a Microsoft Scripting Runtime, pode usar o FileSystemObject diretamente no seu código sem precisar declarar explicitamente a biblioteca.





 Informações sobre o programa

A Ceopag não permite acesso direto ao seu banco de dados, ela disponibiliza arquivos (Querys) gerados no seu painel de vendas 







Iremos precisar de 02 relatórios principais e a partir dele gerar um terceiro que será a base dos nossos cálculos, isto é necessário pois as informações que precisamos não veem adequadamente nos relatórios individuais. 

Primeiro LISTA_VENDAS_TRANSACOES

 Ao ser gerado ele irá informar o PERIODO30122024000000_A_30122024000000_1735642306282.csv

Exemplo

LISTA_VENDAS_TRANSACOES_PERIODO_30122024000000_A_30122024000000_1735642306282.csv

Note que o formato é em .CVS



Relatórios / Minhas Vendas / Presencial Online





Ao abrir, insira o período desejado e clique em Mostrar filtros, escolha APROVADAS

   Clique em Exportar Dados e em seguida escolha Transações o formato em CVS ( Não PDF)

Selecione sua pasta desejada, por exemplo: Relatórios Ceopag



Feito isso, iremos gerar nosso segundo Relatório chamado TRANSAÇÕES_PARCELAS



Clique na opção FINANCEIRO / TRANFERENCIAIS 



  Note que ele também solicita escolher um período , ao fazer isso, ele irá gerar as informações e vai ativar o botão VISUALIZAR (em azul) , clique e salve na mesma pasta por exemplo: Relatórios Ceopag





Observação importante, quando fazemos o controle diários, este relatório é gerado pela Ceopag ficando disponível , caso eles não tenha sido gerado ( raras ocasiões) é necessário solicitar ,  



Pronto.  Nós geramos os dois relatórios necessários para a criação do nosso relatório base principal para cálculos



A configuração VBA acima solicitada, é necessária apenas 1 única vez



Usando, o Sistema de Cálculo para comissões Parceiros MegaPay / Ceopag

Devidamente configurado clique no arquivo 



calculo_comissões.xlsm



ele irá abrir essa tela 































Esse Arquivo é composto por 06 Planilhas

Plan1 (LIsta venda_Transações)

Plan2 (Transações_Parcelas)

Plan3 (vendas_ceopag)

Plan4 (PARCEIROS)

Plan5 (TABELA)

Plan6 (COMISSÃO)



O primeiro passo é importar os dois arquivos gerados 

LISTA_VENDAS_TRANSACOES

TRANSAÇÕES_PARCELAS



Clique no botão Arquivo de Transações , ele irá mostar a Plan1 (LIsta venda_Transações)

Você deverá na pasta aonde o seu arquivo gerado foi saldo e clicar com botão direito e abrir com Excel

Clique na interseção (Quadrado) da linha 1 e coluna A ( isso será sempre padrão)





Já com o arquivo aberto, clique na interseção citada 





Ao clicar neste quadrado (interseção) entre a linha1 e a coluna A

Ele selecionará toda o conteúdo da Planilha

  feito isso, copie para a área de transferência (Ctrl + C)



Agora no programa na Planilha Plan1 (LIsta venda_Transações)

Na mesma interseção clique e cole (Ctrl + V) , pronto, agora repita o processo para o outro arquivo, que será colado na 

 Plan2 (Transações_Parcelas), salve , pronto este foi o seu maior trabalho, agora será super fácil e rápido



GERANDO O RELATÓRIO VENDAS CEOPAG , VÁ PARA PLANILHA 

Plan3 (vendas_ceopag) E CLIQUE NO BOTÃO  GERAR RELATÓRIO, aguarde até a mensagem informar que o relatório foi gerado, feito isso, salve ,

Obs: O botão Limpar , não é necessário clicar, ele vai limpar as 3 planilhas, simultaneamente, isso é usando em casos extremos, pois quando colamos ele mesmo limpa e cola



Agora temo o relatório que precisamos 



Na tela que será aberta automaticamente 

























Perceba que ele vai mostrar uma lista contendo 3 informações



O número do terminal , o nome do parceiro e a tabela que ele está cadastrado de comissionamento, cada tabela possui valores diferentes de percentuais de comissionamento



Ele vai listar  SOMENTE os terminais que tem venda registrada no relatório, se não aparecer o nome do parceiro é pq vc ainda não cadastrou seu parceiro na Plan4 (PARCEIROS), varemos com fazer isso mais adiante



Lógia do processo



Um parceiro pode ter 1 ou diversos clientes, cada cliente pode possui 1 ou diversos terminais, logo

Na Planilha na Plan4 (PARCEIROS), você deve informar TODOS os terminais vinculados ao seu parceiro, cada terminal é um cliente ou o mesmo cliente dele pode ter mais de um terminal, aqui isso não importa. 



Perceba que você poderá SELECIONAR 1 OU DIVERSOS terminais do seu parceiro e ao fazer isso, clique 

CALCULAR SELECIONADOS ( SELECIONE TODOS OU ALGUNS OU UM DO MESMO PARCEIRO DE CADA VEZ) , não pode ser mais de uma parceiro de cada vez



Feito isso ele irá calcular todos os selecionados na Plan6 (COMISSÃO)

Exemplo fictício 





No Exemplo eu selecionei apenas 02 terminais do Parceiro Marco Roberto, ele informa que foram selecionados 2 , realiza todos os cálculos e informa o valor total da comissão baseada nos percentuais de comissões da tabela dele cadastrada, 

Ao clicar no botão ok, ele irá mostrar a Plan6 (COMISSÃO)





































Perceba que ele inseri em cada coluna respectivamente o número do terminal que teve vendas registradas, aqui aparece apenas os terminais que tiveram vendas, ele mostra os totais por cada modalidade e calcula os respectivos valores percentuais em relação a tabela ao qual ele pertence e depois mostra o valor total final ,  



Neste exemplo eu selecionei apenas 2 de 3 , para exemplificação, agora eu irei selecionar os 3 



Clique no botão da maquinha para voltar para a tela principal 





Perceba que agora foram selecionados os 3 terminais pertencentes ao parceiro Marco Roberto, e ao clicar no BOTÃO CALCULAR SELECIONADOS , o resultado já foi diferente



perceba que ele automaticamente preenche as colunas na sequência e refaz os cálculos corretamente

agora iremos fazer com o outro parceiro de exemplo Lucas silva Marques , 



























































Perceba que ele refaz o relatório já com os novos terminais que foram selecionados, 



São as informações necessárias para o pagamento das comissões. 



Na tela principal clique em imprimir, e na caixa de diálogo, vc pode selecionar em PDF, para envio aos seus Parceiros

Enviar via WhatsApp Web por exemplo, 



Como cadastrar os percentuais de comissões (TABELAS)



Na tela principal clique em CADASTRAR TABELAS

Você será direcionado para a Plan5 (TABELA)











































Aqui você irá cadastrar os percentuais de ganho por modalidade para cada tabela, eles estão programas para serrem identificadas por letras Maísculas A , B , C, D, E, F etc....



Cada tabela pode ser associada para cada Parceiro ou para vários Parceiros ao mesmo tempo



Aqui para exemplificação foram cadastrados 3 tabelas (A , B, C) 

É daqui que as taxas são carregadas na Plan Comissão para cada Parceiro selecionado

Os campos em Cinza não devem ser renomeados ou excluídos, vc só poderá mudar os percentuais para cada tabela ou criar outras seguindo a mesma lógica das letras sequenciais. 



CADASTRANDO PARCEIROS



Na tela principal clique em CADASTRAR PARCEIROS

Será apresentada a Plan4 (PARCEIROS)



























Aqui você poderá castrar seu Parceiro , bem como seus terminais para cada parceiro, 

Coluna A = Terminal 

Coluna B = Nome do parceiro

Coluna C = CPF do Parceiro ou CNPJ

Coluna D = Tabela (identificação pela letra Maiúscula)

Sempre que houver vendas registradas para cada terminal, ele irá aparecer na Tela 





Considerações finais



NA PLANILHA, NÃO RENOMEI NADA, PLANILHA 

NÃO EXCLUA OU INSIRA LINHAS OU COLUNAS 

NÃO OCULTE NENHUMA PLANILHA

NÃO MANTENHA FILTROS ATIVOS NAS PLANILHAS 1,2,3

















