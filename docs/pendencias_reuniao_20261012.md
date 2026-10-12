# Dito! — revisão das demandas da reunião (12/10/2026)

Referência do trabalho: branch `refactor/organizacao-base`.
Este arquivo é um acompanhamento do **código**, não atestado de funcionamento em produção.

## Código presente antes do pacote

- [x] Filtro Volume/Ano na listagem de imagens (`core/views.py`, `imagens_lista.html`)
- [x] Exemplo de lote `9_p013mat_cap1` (`organizar_lotes.html`)
- [x] Correção de plural de imagens (`lotes_lista.html`, `imagens_lista.html`)
- [x] Ocultar lotes inativos por padrão na visualização de lotes
- [x] Seed de componentes criado (não executar no banco sem revisar equivalências)
- [x] Serviços de SharePoint e Excel separados inicialmente
- [x] Normalizador FotoWeb e auditoria de importação (testes sem banco)

## Preparado pelo pacote (pendente validação funcional)

- [ ] Importação via **interface**: status FotoWeb respeitado para novos registros; excluídos e legados ignorados; desconhecidos sinalizados; destinos ativos verificados antes da escrita
- [ ] Aviso na descrição quando a origem G27 está explicitamente identificada
- [ ] Notificação para administradores com solicitações de acesso pendentes no dashboard

## Pendente / bloqueado por dependências

- [ ] Importador antigo de linha de comando: requer revisão da lógica de atualização e tipos de campos; **não utilizar** com dados reais até validar
- [ ] Redesign da tela de entrada: falta definir o visual desejado
- [ ] Preferências persistentes por usuário: definir quais preferências e criar migration, com teste em banco separado
- [ ] `seed_status.py` atual altera status já existentes: revisar antes de executar
- [ ] `ALLOWED_HOSTS=['*']` precisa ser restrito para produção (depende de domínios autorizados)
- [ ] Confirmar `DATABASE_URL` persistente e armazenamento de arquivos de mídia na hospedagem
- [ ] Auditar migration `0025` antes de aplicar em banco com dados reais
- [ ] Teste end-to-end com planilha FotoWeb, ambiente de homologação e perfis de usuário
- [ ] Confirmar identificação de imagens reaproveitadas do G27 quando o metadado não explicita a origem

## Observação operacional

Este pacote não executa migrations, não acessa o banco, não altera templates de login e não publica código automaticamente.
