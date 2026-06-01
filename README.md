# Conferência de Folha

Plataforma local para subir a folha de pagamento, a planilha de lançamentos, o portal de empréstimos e o espelho de ponto.

## Como executar

```powershell
& "C:\Users\felip\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" app\server.py
```

Depois abra:

```text
http://127.0.0.1:8878
```

## O que a primeira versão confere

- Lançamentos da planilha contra os demonstrativos da folha.
- Empréstimos do portal contra a rubrica de Crédito do Trabalhador na folha.
- Comparação de líquido quando a folha anterior também é enviada.
- Contagem e período identificado no espelho de ponto.
