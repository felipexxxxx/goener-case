# Validação da entrega 1.1.0

- Ambiente: Windows, Python 3.12.14.
- Qualidade: Ruff sem apontamentos; 61 testes passaram, incluindo OCR dos documentos do case; cobertura total de 86%.
- Entradas: 17 documentos, com uma cópia idêntica registrada por SHA-256 e preservada na planilha.
- Resultado esperado e validado: 12 ok, 4 revisar, 1 falhou. Saldo aprovado pelas regras: R$ 4.114.372,29.
- Os dados originais e os oito campos do modelo foram preservados. Não houve ajuste de saldo para fechar conciliação.
- IA: APIs oficiais OpenAI e Gemini implementadas. Contrato JSON, respostas parciais, bloqueios e cota Gemini testados com HTTP simulado. Nenhuma chamada real foi declarada como validada sem chave. IA desligada nesta entrega de referência.
- Modo consultivo registra recomendação; automático retém saldos quando a análise falha ou recomenda revisão. A IA não sobrepõe as verificações obrigatórias.
- Navegador ChatGPT experimental não homologado: teste parou na confirmação humana antes de enviar mensagem. Não integra o caminho padrão do case.
- HTML conferido em desktop e celular, com busca, filtros e valores em português.
- Estimativa apenas em horas: hipótese de 8 minutos por arquivo e 3 minutos por pendência. Não representa medição da equipe e não converte tempo em dinheiro.
- Power BI adiado. Não houve implantação em nuvem nem envio do case ao portal.

O resumo em `resultado/resumo.json` identifica a execução, duração medida, cache e versão do código. Os testes locais não substituem a conclusão da integração contínua remota.
