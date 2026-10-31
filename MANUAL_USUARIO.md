# Manual do ensalamento ABC Aprendiz

## Para que serve

O aplicativo distribui alunos entre turmas da manhã e da tarde, respeitando a disponibilidade e as vagas. As turmas de um mesmo turno diferem em no máximo um aluno. O modelo de programação inteira equilibra a proporção de meninas e meninos, e a tela permite conferir a solução antes do download.

## 1. Prepare a planilha

O arquivo deve ser `.xlsx` e conter:

| Aba | Colunas | Conteúdo |
| --- | --- | --- |
| `Alunos` | `Nome`, `Sexo`, `Disponibilidade` | Um aluno por linha; sexo `F` ou `M`; disponibilidade `Manhã` ou `Tarde`. |
| `Turmas` | `Recursos`, `Quantidade` | `turmas_manha`, `turmas_tarde`, `vagas_manha`, `vagas_tarde`. As vagas são por turma. |

O programa reconhece diferenças simples de escrita, como `manha`, espaços extras e `feminino`/`masculino`. Essas padronizações aparecem como avisos. Nomes iguais são avisados, mas continuam sendo tratados como alunos distintos.

## 2. Confira a entrada

Abra o aplicativo e selecione a planilha. A prévia mostra as primeiras linhas de `Alunos` e `Turmas`, além do número de alunos por turno e da capacidade informada.

## 3. Escolha quantas turmas ocupar

Os campos **Turmas ocupadas · Manhã** e **Turmas ocupadas · Tarde** aceitam um número de zero até o total disponível. Para usar todas, escolha o total disponível em cada turno. Toda turma escolhida precisa receber pelo menos um aluno. A escolha também precisa oferecer vagas suficientes para todos os alunos do turno.

Se a planilha já tiver `turmas_usadas_manha` e/ou `turmas_usadas_tarde`, esses números aparecem inicialmente na tela. Sem essa informação, o programa propõe o maior número possível de turmas não vazias.

## 4. Gere e confira o resultado

Clique em **Gerar ensalamento**. A tela apresenta totais, gráfico por turma, tabela da distribuição, busca por nome ou turma e uma aba de inconsistências. A otimização pode levar alguns segundos.

Se houver dados inválidos, o aplicativo não gera uma alocação parcial. Ele reúne os avisos e permite baixar uma planilha com a aba `Inconsistências`. Corrija a entrada e envie novamente.

## 5. Baixe os arquivos

- **Excel completo (.xlsx):** cópia da planilha enviada, com as abas `Ensalamento` e `Inconsistências`. A ordem das colunas de `Ensalamento` é `Nome`, `Sexo`, `Disponibilidade`, `Turma`.
- **Resumo por turma (.csv):** quantidade total, meninas, meninos e porcentagem de meninas por turma.
- **Manual (.pdf ou .md):** este documento, disponível na aba **Manual e ajuda** para leitura e download.

O arquivo original nunca é alterado pelo aplicativo.

## Erros comuns

| Aviso | Como resolver |
| --- | --- |
| Falta a aba `Alunos` ou `Turmas` | Confira os nomes das abas. |
| Sexo ou disponibilidade inválidos | Preencha com `F`/`M` e `Manhã`/`Tarde`. |
| Turmas demais para o número de alunos | Reduza o número de turmas ocupadas ou corrija a lista de alunos. |
| Falta de vagas | Aumente o número de turmas ocupadas ou a capacidade informada. |
| Nome duplicado | Verifique o cadastro; se forem pessoas distintas, mantenha as duas linhas. |

## Dados e execução

O processamento acontece localmente no computador em que o Streamlit está rodando. A planilha enviada fica apenas na sessão da aplicação e é usada para montar o arquivo de download. Para iniciar a versão em código, use `streamlit run streamlit_app.py` no ambiente Python 3.12.9 com as dependências de `requirements.txt` instaladas.
