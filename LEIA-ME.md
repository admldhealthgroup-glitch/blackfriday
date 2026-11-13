# LD Health — Catálogo de liquidação

## Instalação
Mantenha `index.html`, `styles.css`, `script.js` e `logo_LD.png` na raiz e `tabela_precos.xlsx` dentro de `dados/`. Publique em hospedagem estática (GitHub Pages, por exemplo). A leitura do Excel via `fetch` exige servidor HTTP/HTTPS; não abra o HTML diretamente com `file://`.

## WhatsApp
No início de `script.js`, substitua `const WHATSAPP_LOJA = "";` pelo número da empresa com DDI e DDD, somente dígitos (ex.: `5511999999999` é fictício). O botão abre o WhatsApp com a mensagem pronta; o cliente precisa enviá-la.

## Preços
O pré-venda utiliza a coluna `PREÇO` da planilha como preço unitário, sem aplicar automaticamente PIX -2% ou desconto por quantidade. Confirme valores e estoque com o cliente. Para alterar a promoção, atualize a planilha antes de publicar.

## Limitações
Sem banco de dados: não há estoque ou reservas em tempo real, histórico centralizado de pedidos ou envio automático de mensagens. O carrinho fica no navegador do cliente (`localStorage`).

## Atualização
A página mostra somente **Preço normal**, sem as outras modalidades de preço. A exportação de PDF foi removida. Todas as referências visíveis a pré-pedido foram alteradas para pré-venda.
