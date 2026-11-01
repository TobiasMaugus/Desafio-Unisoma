# Manual do usuário — ABC Aprendiz

Organização de turmas · UniSoma 2026

## 1. Conheça o aplicativo

O aplicativo organiza alunos em turmas da manhã e da tarde. Ele respeita o turno de cada aluno e as vagas disponíveis, mantém a diferença de tamanho entre turmas ocupadas do mesmo turno em até um aluno e busca equilibrar a proporção de meninas e meninos.

O caminho é: **selecionar a planilha → conferir e salvar os dados → gerar o ensalamento → escolher as salas → conferir e baixar o resultado**. Ajustes manuais são opcionais.

1. Abra o aplicativo. Na versão executável, dê dois cliques em **ABC_Aprendiz.exe**, aguarde a tela de carregamento e a abertura do navegador.
2. Use a aba **Ensalamento** para trabalhar.
3. Consulte **Manual e ajuda** quando precisar. Para ler enquanto trabalha, abra a barra lateral e ative **Abrir manual ao lado**.
4. Para guardar estas instruções ou imprimir, clique em **Baixar manual (.pdf)**. Na barra lateral, use **Baixar manual completo (.pdf)**.

O processamento ocorre no computador que executa o aplicativo. O arquivo original não é sobrescrito: os downloads são novas cópias. Baixe seu trabalho antes de encerrar a sessão ou trocar a planilha.

## 2. Prepare e selecione a planilha

Use um arquivo **.xlsx** com as abas **Alunos** e **Turmas**, com os nomes abaixo na primeira linha. Mantenha a estrutura da planilha.

### Aba Alunos

- **Nome:** nome do aluno, um aluno por linha.
- **Sexo:** preencha com **F** ou **M**.
- **Disponibilidade:** preencha com **Manhã** ou **Tarde**.

Use valores preenchidos diretamente nas células de cadastro. Fórmulas nesses campos precisam ser substituídas por valores. Linhas totalmente vazias são ignoradas. Nomes repetidos geram aviso; se forem pessoas diferentes, mantenha ambos os registros.

### Aba Turmas

As colunas são **Recursos** e **Quantidade**. Preencha os quatro recursos:

- **turmas_manha:** quantidade de salas disponíveis pela manhã.
- **turmas_tarde:** quantidade de salas disponíveis à tarde.
- **vagas_manha:** vagas por sala da manhã, e não a soma de vagas do turno.
- **vagas_tarde:** vagas por sala da tarde, e não a soma de vagas do turno.

Exemplo: para três salas da manhã com vinte vagas em cada uma, informe **turmas_manha = 3** e **vagas_manha = 20**. Isso representa sessenta vagas pela manhã.

Em **1. Selecionar planilha**, clique no campo de envio ou arraste o arquivo para ele. Aguarde a exibição das tabelas. Se o arquivo não abrir, confirme que é um .xlsx válido com as abas e os cabeçalhos exigidos.

## 3. Confira os dados e as correções automáticas

Em **2. Conferir dados e gerar o ensalamento**, revise **Turmas e capacidades** e **Alunos**.

1. Confira a quantidade de salas e as vagas por sala em cada turno.
2. Confira nome, sexo e disponibilidade dos alunos.
3. Leia a coluna **Atenção**. Vermelho indica erro; amarelo, aviso; verde, correção automática. O símbolo **—** indica ausência de apontamento naquele registro.
4. Para ampliar a tabela de alunos, use o botão de **tela cheia** na barra da tabela. Use o mesmo botão para voltar. Se necessário, deslize a barra horizontal até **Atenção** e aumente a largura da coluna arrastando a borda do cabeçalho.

Os avisos dos alunos usam o número da coluna **Registro**. Esse número identifica o cadastro, mesmo após a exclusão de outras linhas; ele não é a posição atual da linha na tela. Os apontamentos são atualizados conforme as edições.

O aplicativo padroniza algumas formas de escrita, por exemplo **feminino → F**, **masculino → M** e **manha → Manhã**. As normalizações reconhecidas ficam indicadas como correções automáticas, com os valores originais e corrigidos no relatório. Valores desconhecidos, fórmulas e caracteres invisíveis apontados como erro exigem revisão.

## 4. Edite, adicione ou exclua dados

### Corrigir alunos

1. Clique na célula de **Nome**, **Sexo** ou **Disponibilidade** e digite o valor correto. Finalize com Enter ou clique fora da célula.
2. Para adicionar um aluno, clique no **+** no rodapé da coluna **Atenção** e preencha os campos da nova linha.
3. Para excluir alunos do cadastro, marque as caixas das linhas desejadas e clique no botão com a **lixeira** acima da tabela.
4. Confira os apontamentos e clique em **Salvar alunos**. Para abandonar as edições ainda não salvas, clique em **Descartar alterações** junto à tabela de alunos.

**Registro** e **Atenção** são preenchidos pelo aplicativo e não podem ser editados. Excluir um cadastro retira esse aluno da entrada que será usada no próximo cálculo.

### Corrigir turmas e capacidades

1. Edite as configurações ou quantidades necessárias na tabela **Turmas e capacidades**.
2. Clique em **Salvar turmas**. Para abandonar essas edições ainda não salvas, use **Descartar alterações** junto à tabela de turmas.

### Desfazer e restaurar dados

Abra **Histórico e restauração** abaixo da tabela correspondente. Use **Desfazer** ou **Refazer** para percorrer alterações salvas. Para voltar aos dados importados, marque a confirmação e clique em **Restaurar dados originais**. Revise novamente as normalizações e os avisos.

Enquanto houver edições não salvas, o resultado anterior fica oculto. Salvar dados ou restaurar o cadastro exige gerar o ensalamento novamente e reinicia a escolha de salas. O botão **Baixar planilha atualizada** exporta os dados salvos e normalizados; salve as edições antes de usá-lo.

## 5. Gere o primeiro ensalamento

1. Corrija os erros vermelhos e salve as alterações nas duas tabelas.
2. Clique em **Gerar ensalamento** e aguarde a conclusão.
3. Leia as mensagens e o **Relatório geral de erros e avisos**, quando disponível.
4. Em **3. Consultar e baixar o resultado**, confira os totais e o resumo por turma.

A primeira distribuição usa **Todas as salas disponíveis**. Em turno sem alunos, não há turma ocupada. Se houver alunos, mas menos alunos do que salas na configuração inicial, confira o relatório e ajuste a quantidade de salas disponíveis na tabela antes de gerar novamente.

Erros bloqueantes impedem a geração. Nesse caso, use **Gerar relatório de inconsistências**, quando exibido, e **Baixar planilha com inconsistências** para consultar os problemas. Corrija os dados, salve e gere novamente. Avisos não bloqueantes, como nomes repetidos, permitem gerar o ensalamento, mas precisam ser conferidos.

## 6. Escolha como usar as salas

Após gerar o resultado, localize **Quer usar menos salas?**. Há duas opções:

- **Todas as salas disponíveis:** utiliza as salas disponíveis de cada turno com alunos.
- **Somente as salas necessárias:** utiliza o menor número de salas que comporta os alunos de cada turno, dentro das salas disponíveis.

Para mudar a distribuição:

1. Clique em **Usar somente as salas necessárias** ou **Usar todas as salas**. O card escolhido mostra **Selecionado**.
2. Confira a **Prévia da configuração por turno**, com salas usadas, vagas por sala e vagas livres.
3. Clique em **Recalcular ensalamento**.
4. Aguarde o novo resultado e confira as turmas antes de baixar.

Escolher o card apenas altera a configuração pretendida. Enquanto o recálculo estiver pendente, a tela avisa que o resultado é anterior e os downloads da distribuição ficam indisponíveis. Após recalcular, os arquivos correspondem à opção escolhida. Recalcular também reinicia os ajustes manuais do resultado anterior.

Nas duas opções, o aplicativo respeita as vagas e mantém a diferença entre turmas ocupadas do mesmo turno em até um aluno. Escolher menos salas não cria vagas adicionais.

## 7. Consulte e baixe o resultado

Confira **Alunos alocados**, **Turmas ocupadas** e **Resumo por turma**. O resumo informa alunos, vagas, vagas restantes, meninas, meninos e percentuais. Abra **Ver gráfico por turma** para visualizar a distribuição.

Na aba **Lista de alunos**, use **Buscar aluno** para localizar um nome ou parte dele e **Turno** para filtrar manhã ou tarde. Esses filtros servem para consulta; não retiram alunos do ensalamento exportado.

- **Baixar Excel completo (.xlsx):** exporta o resultado otimizado vigente, com as abas de entrada, **Ensalamento** e **Inconsistências**. A aba Ensalamento contém **Nome**, **Sexo**, **Disponibilidade** e **Turma**.
- **Outros formatos de download → Baixar resumo por turma (.csv):** exporta os totais e percentuais por turma.
- **Outros formatos de download → Baixar resumo da distribuição (PDF):** exporta o resumo, os apontamentos e as listas de alunos separadas por turma, com continuação em outras páginas quando necessário.
- **Baixar planilha atualizada:** na seção de conferência, exporta o cadastro salvo para uso posterior. Esse arquivo não substitui o download do resultado do ensalamento.

Os downloads do resultado otimizado continuam representando a solução calculada. Para levar alterações feitas em **Ajustes avançados**, use os botões específicos de resultado ajustado descritos a seguir.

## 8. Faça ajustes manuais somente se necessário

Na aba **Ajustes avançados**, leia o aviso vermelho: **Use esta área somente se precisar alterar a solução otimizada. As mudanças podem afetar o equilíbrio entre turmas.**

Confira os totais por turno e use **Consultar turma** para ver seus alunos. Os ajustes afetam uma cópia da distribuição, preservando a solução otimizada original. Eles não alteram a tabela de cadastro de entrada.

### Trocar dois alunos

1. Abra **Trocar alunos entre turmas** e escolha as turmas e os alunos.
2. Confira a prévia e os totais antes e depois.
3. As turmas devem ser do mesmo turno. Se a troca envolver sexos diferentes, leia o aviso e marque a confirmação solicitada.
4. Clique em **Confirmar troca**.

### Mover um aluno

1. Abra **Mover aluno**, selecione o aluno e a **Turma de destino**.
2. A turma deve ser do mesmo turno e ter vaga disponível.
3. Confira a prévia, os totais e o aviso de equilíbrio. Marque **Confirmo esta movimentação manual.**
4. Clique em **Confirmar movimentação**.

### Remover um aluno apenas da distribuição ajustada

1. Abra **Remover aluno do ensalamento** e escolha a turma e o aluno.
2. Leia o aviso e marque **Confirmo que este aluno ficará sem turma.**
3. Clique em **Remover aluno**. O aluno permanece no cadastro, mas fica sem turma no resultado ajustado.

### Desfazer, restaurar e baixar os ajustes

- Use **Desfazer última alteração** para voltar uma etapa dos ajustes manuais.
- Para voltar à distribuição calculada, marque **Confirmo restaurar a solução original.** e clique em **Restaurar ensalamento original**.
- Use **Baixar ensalamento ajustado** para exportar o Excel com os ajustes.
- Após uma alteração, use **Baixar relatório da distribuição ajustada (PDF)** para exportar o relatório correspondente. Alunos removidos aparecem na lista de alunos sem turma.

Confira todos os avisos antes de usar o arquivo. Se o ajuste deixar problemas, como aluno sem turma ou desequilíbrio de quantidade, a exportação Excel será identificada como **NAO_CONFORME** e os problemas constarão na aba **Inconsistências**. Situações que excedam a capacidade ou envolvam outro turno impedem o download ajustado. Ajustar manualmente não executa uma nova otimização.

## 9. Resolva os problemas mais comuns

- **Arquivo não abre ou falta uma aba:** confira o formato .xlsx, as abas Alunos e Turmas e seus cabeçalhos.
- **Sexo ou disponibilidade inválidos:** corrija para F/M e Manhã/Tarde e salve os alunos.
- **Fórmula ou caractere invisível:** substitua o conteúdo indicado por um valor simples digitado diretamente na célula.
- **Nome duplicado:** confirme se são pessoas distintas. Corrija ou exclua o cadastro repetido apenas se necessário.
- **Salas demais para o número de alunos na primeira geração:** revise a quantidade disponível por turno na tabela, salve e gere novamente.
- **Falta de vagas:** confira a quantidade de salas e as vagas reais por sala; corrija o cadastro conforme a disponibilidade real.
- **Resultado oculto:** há edições pendentes. Salve ou descarte as alterações na tabela correspondente.
- **Download indisponível após mudar as salas:** clique em Recalcular ensalamento e aguarde a conclusão.
- **Falha no cálculo:** confira os dados e a mensagem apresentada. Se persistir, guarde a planilha e a mensagem para solicitar suporte.

Antes de encerrar, baixe a planilha atualizada se editou o cadastro, o resultado da opção de salas desejada e, se necessário, o resultado ajustado. Guarde os arquivos com nomes que permitam reconhecer qual versão será utilizada.
