
"use strict";

// LD HEALTH INTERNATIONAL
// Arquivo: script.js

const ARQUIVO_EXCEL = "dados/tabela_precos.xlsx";
const WHATSAPP_LOJA = "5512988214515";

const nomesCategorias = {
  toxinas: "TOXINAS BOTULÍNICAS",
  preenchedores: "PREENCHEDORES",
  bioestimuladores: "BIOESTIMULADORES",
  fios: "FIOS DE PDO",
  microcanulas: "MICROCÂNULAS",
  solucoes: "SOLUÇÕES E MEDICAMENTOS AUXILIARES",
  dermapen: "DERMAPEN E CARTUCHOS",
  skincare: "SKINCARE E BOOSTERS",
  insumos: "INSUMOS",
};

const $ = (id) => document.getElementById(id);
const produtos = Object.fromEntries(
  Object.keys(nomesCategorias).map((k) => [k, []]),
);

const CHAVE_CARRINHO = "ldhealth_pedido_v2";
let carrinho = [];
let carregado = false;

function normalizar(texto) {
  return String(texto ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function nomeBonito(texto) {
  const s = String(texto ?? "")
    .trim()
    .toLocaleLowerCase("pt-BR");
  return s ? s.charAt(0).toLocaleUpperCase("pt-BR") + s.slice(1) : "";
}

function precoNumero(valor) {
  if (typeof valor === "number") {
    return Number.isFinite(valor) ? valor : null;
  }
  if (valor == null || String(valor).trim() === "") {
    return null;
  }
  let s = String(valor).replace(/R\$/gi, "").replace(/\s/g, "");
  if (s.includes(",")) {
    s = s.replace(/\./g, "").replace(",", ".");
  }
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function moeda(n) {
  return Number(n).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });
}

function status(mensagem, erro = false) {
  const el = $("data-status");
  if (!el) return;
  el.textContent = mensagem;
  el.classList.toggle("error", erro);
  el.hidden = !mensagem;
}

function chaveCategoria(texto) {
  const s = normalizar(texto);
  const mapa = {
    "TOXINAS BOTULINICAS": "toxinas",
    PREENCHEDORES: "preenchedores",
    BIOESTIMULADORES: "bioestimuladores",
    "FIOS DE PDO": "fios",
    "FIOS PDO": "fios",
    MICROCANULAS: "microcanulas",
    "SOLUCOES E MEDICAMENTOS AUXILIARES": "solucoes",
    "DERMAPEN E CARTUCHOS": "dermapen",
    "SKINCARE E BOOSTERS": "skincare",
    INSUMOS: "insumos",
  };
  return mapa[s] || null;
}

// CARREGAMENTO DO EXCEL

async function carregarExcel() {
  try {
    status("Carregando produtos...");

    if (typeof XLSX === "undefined") {
      throw new Error("Biblioteca Excel indisponível.");
    }

    const resposta = await fetch(ARQUIVO_EXCEL + "?v=" + Date.now(), {
      cache: "no-store",
    });

    if (!resposta.ok) {
      throw new Error("Planilha não encontrada em dados/tabela_precos.xlsx");
    }

    const bytes = await resposta.arrayBuffer();
    const livro = XLSX.read(bytes, { type: "array" });
    const aba = livro.Sheets[livro.SheetNames[0]];

    if (!aba) throw new Error("Planilha sem abas.");

    const linhas = XLSX.utils.sheet_to_json(aba, {
      header: 1,
      defval: "",
      raw: true,
    });

    if (!linhas.length) throw new Error("Planilha vazia.");

    const cab = linhas[0].map(normalizar);
    const cat = cab.indexOf("CATEGORIA");
    const desc = cab.findIndex((x) =>
      ["DESCRICAO", "PRODUTO", "PRODUTOS"].includes(x),
    );
    const val = cab.findIndex((x) => ["PRECO", "VALOR"].includes(x));

    if ([cat, desc, val].includes(-1)) {
      throw new Error(
        "A primeira linha precisa ter CATEGORIA, DESCRIÇÃO e PREÇO.",
      );
    }

    const novaBase = Object.fromEntries(
      Object.keys(nomesCategorias).map((k) => [k, []]),
    );

    linhas.slice(1).forEach((linha, i) => {
      if (linha.every((v) => String(v ?? "").trim() === "")) return;

      const categoria = chaveCategoria(linha[cat]);
      if (!categoria) {
        throw new Error("Categoria inválida na linha " + (i + 2));
      }

      const descricao = String(linha[desc] ?? "").trim();
      const preco = precoNumero(linha[val]);

      if (!descricao || preco === null || preco < 0) {
        throw new Error("Produto ou preço inválido na linha " + (i + 2));
      }

      novaBase[categoria].push({
        id: categoria + "-" + (i + 2),
        categoria,
        descricao,
        preco,
      });
    });

    Object.keys(produtos).forEach((k) => {
      produtos[k] = novaBase[k];
    });

    carregado = true;
    status("");
    pesquisar();
  } catch (erro) {
    carregado = false;
    status("Erro: " + erro.message, true);
    console.error(erro);
  }
}

// CARRINHO

function recuperarCarrinho() {
  try {
    const dados = JSON.parse(localStorage.getItem(CHAVE_CARRINHO) || "[]");
    if (Array.isArray(dados)) {
      carrinho = dados.filter(
        (x) =>
          x &&
          typeof x.descricao === "string" &&
          Number.isFinite(x.preco) &&
          Number.isInteger(x.quantidade) &&
          x.quantidade > 0 &&
          x.quantidade <= 999,
      );
    }
  } catch (e) {
    carrinho = [];
  }
}

function quantidadeProduto(produto) {
  return carrinho.find((x) => x.id === produto.id)?.quantidade || 0;
}

function alterarQuantidade(produto, diferenca) {
  const atual = carrinho.find((x) => x.id === produto.id);
  const quantidade = (atual?.quantidade || 0) + diferenca;

  if (quantidade < 0 || quantidade > 999) return;

  if (quantidade === 0) {
    carrinho = carrinho.filter((x) => x.id !== produto.id);
  } else if (atual) {
    atual.quantidade = quantidade;
    atual.preco = produto.preco;
  } else {
    carrinho.push({
      id: produto.id,
      descricao: produto.descricao,
      preco: produto.preco,
      quantidade,
    });
  }

  localStorage.setItem(CHAVE_CARRINHO, JSON.stringify(carrinho));
  renderizarCarrinho();
  atualizarQuantidades();
}

function criarControles(produto) {
  const div = document.createElement("div");
  div.className = "inline-purchase-controls";
  div.dataset.produtoId = produto.id;

  const menos = document.createElement("button");
  menos.type = "button";
  menos.className = "quantity-button";
  menos.textContent = "−";
  menos.setAttribute("aria-label", "Diminuir " + produto.descricao);

  const quantidade = document.createElement("output");
  quantidade.className = "quantity-value";

  const mais = document.createElement("button");
  mais.type = "button";
  mais.className = "quantity-button";
  mais.textContent = "+";
  mais.setAttribute("aria-label", "Aumentar " + produto.descricao);

  menos.onclick = () => alterarQuantidade(produto, -1);
  mais.onclick = () => alterarQuantidade(produto, 1);

  div.append(menos, quantidade, mais);
  div.atualizar = () => {
    const q = quantidadeProduto(produto);
    quantidade.textContent = String(q);
    menos.disabled = q === 0;
    mais.disabled = q >= 999;
  };
  div.atualizar();
  return div;
}

function atualizarQuantidades() {
  document
    .querySelectorAll(".inline-purchase-controls")
    .forEach((el) => el.atualizar?.());
}

function renderizarCarrinho() {
  const area = $("cart-items");
  if (!area) return;
  area.replaceChildren();

  const total = carrinho.reduce((s, x) => s + x.preco * x.quantidade, 0);

  $("cart-empty").hidden = carrinho.length > 0;
  $("cart-summary").hidden = carrinho.length === 0;
  $("cart-total").textContent = moeda(total);
  $("send-whatsapp").disabled = carrinho.length === 0;

  carrinho.forEach((item) => {
    const linha = document.createElement("div");
    linha.className = "cart-item";

    const info = document.createElement("div");
    info.className = "cart-item-info";

    const titulo = document.createElement("strong");
    titulo.textContent = nomeBonito(item.descricao);

    const valor = document.createElement("span");
    valor.textContent =
      item.quantidade +
      " × " +
      moeda(item.preco) +
      " = " +
      moeda(item.quantidade * item.preco);

    info.append(titulo, valor);
    linha.append(info, criarControles(item));
    area.append(linha);
  });
}

// CATEGORIAS

function abrirCategoria(categoria) {
  if (!nomesCategorias[categoria]) return;

  $("category-title").textContent = nomesCategorias[categoria];
  const corpo = $("products-list");
  corpo.replaceChildren();

  const lista = produtos[categoria] || [];
  $("empty-message").hidden = lista.length > 0;

  lista.forEach((produto) => {
    const tr = document.createElement("tr");
    tr.className = "product-purchase-row";

    const tdNome = document.createElement("td");
    tdNome.textContent = nomeBonito(produto.descricao);

    const tdPreco = document.createElement("td");
    tdPreco.textContent = moeda(produto.preco);

    const tdQtd = document.createElement("td");
    tdQtd.append(criarControles(produto));

    tr.append(tdNome, tdPreco, tdQtd);
    corpo.append(tr);
  });

  $("categories-view").hidden = true;
  $("products-view").hidden = false;
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function voltarCategorias() {
  $("products-view").hidden = true;
  $("categories-view").hidden = false;
  window.scrollTo({ top: 0, behavior: "smooth" });
}

// PESQUISA

function pesquisar() {
  const campo = $("product-search");
  const area = $("search-results");
  const termo = normalizar(campo.value);

  $("clear-search-button").hidden = !termo;
  area.replaceChildren();
  area.hidden = !termo;

  if (!termo) return;

  if (!carregado) {
    area.textContent = "Aguarde o carregamento dos produtos.";
    return;
  }

  const encontrados = Object.values(produtos)
    .flat()
    .filter((p) => normalizar(p.descricao).includes(termo));

  const titulo = document.createElement("div");
  titulo.className = "search-results-header";
  titulo.textContent = encontrados.length + " produto(s) encontrado(s)";
  area.append(titulo);

  if (!encontrados.length) {
    const vazio = document.createElement("p");
    vazio.textContent = "Nenhum produto encontrado.";
    area.append(vazio);
    return;
  }

  const lista = document.createElement("div");
  lista.className = "search-results-list";

  encontrados.forEach((produto) => {
    const botao = document.createElement("button");
    botao.type = "button";
    botao.className = "search-result";

    const info = document.createElement("div");
    info.className = "search-result-info";

    const nome = document.createElement("span");
    nome.className = "search-result-name";
    nome.textContent = nomeBonito(produto.descricao);

    const categoria = document.createElement("span");
    categoria.className = "search-result-category";
    categoria.textContent = nomesCategorias[produto.categoria];

    const preco = document.createElement("span");
    preco.className = "search-result-price";
    preco.textContent = moeda(produto.preco);

    info.append(nome, categoria);
    botao.append(info, preco);

    botao.onclick = () => abrirCategoria(produto.categoria);
    lista.append(botao);
  });

  area.append(lista);
}

// NOTINHA EM IMAGEM

function linhasTexto(ctx, texto, largura) {
  const palavras = texto.split(/\s+/);
  const linhas = [];
  let atual = "";

  palavras.forEach((p) => {
    const teste = atual ? atual + " " + p : p;
    if (atual && ctx.measureText(teste).width > largura) {
      linhas.push(atual);
      atual = p;
    } else {
      atual = teste;
    }
  });

  if (atual) linhas.push(atual);
  return linhas;
}

async function carregarLogo() {
  const img = new Image();
  img.src = "logo_LD.png";
  await img.decode();
  return img;
}

async function gerarNotinha(cliente) {
  const largura = 600;
  const escala = 2;
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");

  if (!ctx) throw new Error("Canvas indisponível.");

  ctx.font = "20px Arial";

  const itens = carrinho.map((x) => ({
    ...x,
    linhas: linhasTexto(ctx, nomeBonito(x.descricao), 500),
  }));

  const alturaItens = itens.reduce((s, x) => s + 68 + x.linhas.length * 29, 0);

  const altura = 430 + alturaItens;

  canvas.width = largura * escala;
  canvas.height = altura * escala;
  ctx.scale(escala, escala);

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, largura, altura);

  ctx.fillStyle = "#123c78";
  ctx.fillRect(0, 0, largura, 9);

  const logo = await carregarLogo();
  const fator = Math.min(340 / logo.naturalWidth, 115 / logo.naturalHeight);
  const w = logo.naturalWidth * fator;
  const h = logo.naturalHeight * fator;

  ctx.drawImage(logo, (largura - w) / 2, 25, w, h);

  ctx.textAlign = "center";
  ctx.fillStyle = "#123c78";
  ctx.font = "bold 24px Arial";
  ctx.fillText("RESUMO DO PEDIDO", 300, 175);

  ctx.fillStyle = "#64748b";
  ctx.font = "16px Arial";
  ctx.fillText(new Date().toLocaleDateString("pt-BR"), 300, 205);

  ctx.strokeStyle = "#dce5f0";
  ctx.beginPath();
  ctx.moveTo(40, 225);
  ctx.lineTo(560, 225);
  ctx.stroke();

  ctx.textAlign = "left";
  ctx.fillStyle = "#64748b";
  ctx.font = "15px Arial";
  ctx.fillText("CLIENTE", 40, 255);

  ctx.fillStyle = "#17233d";
  ctx.font = "bold 21px Arial";
  ctx.fillText(cliente, 40, 286, 520);

  ctx.fillStyle = "#64748b";
  ctx.font = "bold 15px Arial";
  ctx.fillText("PRODUTOS", 40, 325);

  let y = 355;

  itens.forEach((item) => {
    ctx.textAlign = "left";
    ctx.fillStyle = "#17233d";
    ctx.font = "20px Arial";

    item.linhas.forEach((linha) => {
      ctx.fillText(linha, 40, y);
      y += 29;
    });

    ctx.fillStyle = "#64748b";
    ctx.font = "16px Arial";
    ctx.fillText(item.quantidade + " × " + moeda(item.preco), 40, y + 8);

    ctx.textAlign = "right";
    ctx.fillStyle = "#123c78";
    ctx.font = "bold 18px Arial";
    ctx.fillText(moeda(item.quantidade * item.preco), 560, y + 8);

    y += 62;

    ctx.strokeStyle = "#e8edf4";
    ctx.beginPath();
    ctx.moveTo(40, y - 20);
    ctx.lineTo(560, y - 20);
    ctx.stroke();
  });

  const total = carrinho.reduce((s, x) => s + x.preco * x.quantidade, 0);

  ctx.fillStyle = "#edf4ff";
  ctx.fillRect(30, y, 540, 75);

  ctx.textAlign = "left";
  ctx.fillStyle = "#123c78";
  ctx.font = "bold 21px Arial";
  ctx.fillText("TOTAL", 48, y + 47);

  ctx.textAlign = "right";
  ctx.font = "bold 27px Arial";
  ctx.fillText(moeda(total), 550, y + 47);

  ctx.textAlign = "center";
  ctx.fillStyle = "#64748b";
  ctx.font = "13px Arial";
  ctx.fillText(
    "Valores e disponibilidade sujeitos à confirmação.",
    300,
    y + 110,
  );

  const blob = await new Promise((resolve) =>
    canvas.toBlob(resolve, "image/png"),
  );

  if (!blob) throw new Error("Não foi possível criar a imagem.");

  return new File([blob], "pedido-ld-health.png", {
    type: "image/png",
  });
}

// COMPARTILHAMENTO

function limparPedido() {
  carrinho = [];

  localStorage.removeItem(CHAVE_CARRINHO);

  renderizarCarrinho();
  atualizarQuantidades();

  const campoNome = document.getElementById("customer-name");
  if (campoNome) {
    campoNome.value = "";
    campoNome.setCustomValidity("");
  }

  console.info("Pedido finalizado. Carrinho limpo.");
}


async function enviarPedido() {
  if (!carrinho.length) return;

  const campo = $("customer-name");
  const nome = campo.value.trim();

  if (nome.length < 8) {
    campo.setCustomValidity("Informe seu nome com pelo menos 8 caracteres.");
    campo.reportValidity();
    campo.focus();
    return;
  }

  campo.setCustomValidity("");

  const total = carrinho.reduce(
    (s, item) => s + item.preco * item.quantidade, 0
  );

  const linhas = [
    "*PEDIDO LD HEALTH INTERNATIONAL*",
    "",
    "*Cliente:* " + nome,
    "*Data:* " + new Date().toLocaleDateString("pt-BR"),
    "",
    "*PRODUTOS*",
    ...carrinho.flatMap((item) => [
      "• " + nomeBonito(item.descricao),
      "  " + item.quantidade + " x " + moeda(item.preco) +
        " = " + moeda(item.quantidade * item.preco),
      ""
    ]),
    "*TOTAL: " + moeda(total) + "*",
    "",
    "Valores e disponibilidade sujeitos à confirmação."
  ];

  const numero = "5512988214515";
  const mensagem = encodeURIComponent(linhas.join("\n"));

  window.location.href =
    "https://wa.me/" + numero + "?text=" + mensagem;
}


// INICIALIZAÇÃO

document.addEventListener("DOMContentLoaded", () => {
  recuperarCarrinho();
  renderizarCarrinho();

  document.querySelectorAll(".category-card").forEach((card) => {
    card.addEventListener("click", () => abrirCategoria(card.dataset.category));
  });

  $("back-button").addEventListener("click", voltarCategorias);

  $("product-search").addEventListener("input", pesquisar);

  $("clear-search-button").addEventListener("click", () => {
    $("product-search").value = "";
    pesquisar();
    $("product-search").focus();
  });

  $("customer-name").addEventListener("input", (e) => {
    e.target.setCustomValidity("");
  });

  $("send-whatsapp").addEventListener("click", enviarPedido);

  carregarExcel();
});

// Abertura promocional: LED -> etiqueta -> catálogo.
window.addEventListener("load", () => {
  const intro = $("intro");
  const main = $("main-content");
  if (!intro || !main) return;

  let concluida = false;
  const finalizar = () => {
    if (concluida) return;
    concluida = true;
    intro.classList.add("hidden");
    main.classList.add("visible");
    window.setTimeout(() => { intro.hidden = true; }, 750);
  };

  $("skip-intro")?.addEventListener("click", finalizar);

  const reduzMovimento = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (reduzMovimento) {
    finalizar();
    return;
  }
  window.setTimeout(() => {
    if (!concluida) intro.classList.add("show-tag");
  }, 2600);
  window.setTimeout(finalizar, 1500);
});
