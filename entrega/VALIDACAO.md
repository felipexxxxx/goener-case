# Validação da entrega

Validação em 14/11/2026, com Python 3.12.14 no Windows. Código do processamento: commit `582d157dedc5680e06fe41ed134f3934ac5f0a89`.

## Execução dos arquivos fornecidos

| Verificação | Resultado observado |
| --- | --- |
| Entrada | 17 arquivos: 7 PDFs, 5 planilhas e 5 imagens. |
| Saída | 17 linhas, além do cabeçalho, no modelo original. |
| Classificação | 12 `ok`, 4 `revisar`, 1 `falhou`. |
| Execução sem cache | 32,661 segundos; saída entregue em `resultado/`. |
| Reexecução com cache preenchido | 0,178 segundo; 16 extrações reaproveitadas. O comunicado inválido é lido novamente. |
| Idempotência | As oito colunas e as 17 linhas ficaram iguais na execução sem cache e nas duas execuções seguintes. |
| Integridade | SHA-256 dos originais preservado. Hashes dos entregáveis conferidos contra o manifesto. |
| Planilha | Valores reabertos e conferidos; sem fórmulas injetadas; abas e colunas do modelo preservadas. |
| Documentação | PDF renderizado e conferido visualmente; uma página. |

## Ressalvas encontradas

- **Cópia de arquivo:** `EXT_102026_LagoaAzul (1).pdf` tem conteúdo idêntico ao arquivo sem sufixo. A linha da cópia permanece registrada e fora das somas de acompanhamento.
- **Possível movimento repetido:** os dois extratos idênticos de Lagoa Azul contêm dois lançamentos semelhantes de R$ 5.830,06 em 26/10/2026. São sinalizados para conferência, sem exclusão automática.
- **Saldo divergente:** Sol do Sertão declara R$ 457.179,12; os movimentos terminam em R$ 456.099,12. Diferença de R$ 1.080,00. O declarado foi preservado.
- **Outra competência:** Serra Dourada II corresponde a setembro de 2026, enquanto o fechamento configurado é outubro.
- **Documento sem extrato:** o aviso da agência é um comunicado. Sua linha recebe `falhou`, com saldo vazio e causa informada.

## Testes

- **49 testes aprovados**, incluindo inferência OCR real das cinco imagens.
- **90% de cobertura de linhas** na suíte local completa.
- Ruff sem ocorrências e formatação conferida.
- Integração contínua aprovada em **Ubuntu e Windows**, incluindo instalação limpa, testes e execução do pipeline a partir da pasta de entrada.
- Comando PowerShell testado, incluindo preparação do ambiente, execução e reexecução.

[Execução do GitHub Actions](https://github.com/felipexxxxx/goener-case/actions/runs/43557008743)

Os tempos são observações desta máquina, sem garantia de desempenho em outros ambientes. Os ganhos de horas da documentação são uma estimativa com premissas explícitas.
