# Decisões de implementação

## Fluxo contínuo e portabilidade

O programa recebe uma pasta, extrai os dados e publica a planilha e o log na mesma execução. Não consulta o chat, não depende do Codex e não usa valores digitados a partir da inspeção dos arquivos. O leitor de Excel também escreve no modelo usando openpyxl, para que o fluxo seja instalável em uma máquina comum com Python. O modelo original fica preservado em `templates/`.

## Regra de negócio

Uma linha por arquivo, inclusive duplicatas e falhas. A duplicata mantém o saldo extraído e recebe `revisar`; não é removida. Os totais de acompanhamento separam valores aprovados de valores retidos. Nenhum saldo é corrigido automaticamente para fazer a conciliação fechar.

## OCR local

PDF com texto usa leitura direta. Imagem e PDF digitalizado usam OCR em CPU, sem chave de API. A confiança é um sinal adicional, não prova da correção do saldo: as movimentações também são conciliadas. Modelos são preparados na instalação; arquivos financeiros não são enviados a provedores.

## Cache e retomada

SQLite guarda extrações por hash do conteúdo e versão do extrator. A validação de competência e a classificação de duplicatas são refeitas a cada execução. Uma falha de publicação não exige repetir OCR. O cache é uma otimização e pode ser desligado para uma execução independente.

## Operação

O lote isola erros de documentos. Erro de entrada, configuração, lock ou publicação impede declarar conclusão. STOP solicita interrupção entre arquivos; os restantes ficam registrados como cancelados. Retries atendem apenas falhas transitórias de leitura, nunca erros de interpretação.

## Etapa de IA opcional

A integração HTTP com a Responses API está implementada com instruções versionadas e JSON validado, mas a chamada real depende de chave e não foi executada sem ela. Permanece desligada por padrão devido à possível cobrança. O documento é tratado como dado, sem autoridade para mudar instruções. Os testes de contrato usam respostas simuladas explicitamente identificadas.

Dois modos foram definidos com o responsável pelo case: consultivo, que registra recomendação e justificativa sem alterar os totais; e automático, que aplica a decisão dentro de uma política aprovada pela gestão. A evidência deve separar resultado das regras, recomendação da IA, decisão aplicada e eventual revisão humana. O modelo não poderá corrigir saldos ou executar código livremente. Falha de API, recusa ou resposta inválida manterá a pendência para revisão.

O desenvolvimento supervisionado começará com exemplos classificados pela equipe, avaliação em casos separados e registro das divergências. Ajustar instruções e exemplos não equivale a treinar os pesos do modelo. Uma eventual etapa de fine-tuning depende de dados suficientes e avaliação específica. A passagem para o modo automático será discutida com a gestão após medir erros e critérios de aceitação.

O pipeline atual continua sequencial. Paralelismo é uma evolução possível por arquivo, com limite de trabalhadores e escrita centralizada da planilha e do checkpoint; ainda não foi implementado nem medido.
