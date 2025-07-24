let assistentes = [];
let mensageiros = { A: [], B: [] };
let diasVisiveis = 7;
let semanaAtual = 0;
let escalaCompleta = [];
let mesAtual = new Date().getMonth();
let anoAtual = new Date().getFullYear();

function preencherSelects() {
  const mesSelect = document.getElementById("mesSelect");
  const anoSelect = document.getElementById("anoSelect");

  mesSelect.innerHTML = "";
  anoSelect.innerHTML = "";

  for (let m = 0; m < 12; m++) {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = new Date(0, m).toLocaleString("pt-BR", { month: "long" });
    if (m === mesAtual) opt.selected = true;
    mesSelect.appendChild(opt);
  }

  const anoReal = new Date().getFullYear();
  for (let y = anoReal - 2; y <= anoReal + 2; y++) {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
    if (y === anoAtual) opt.selected = true;
    anoSelect.appendChild(opt);
  }
}

function carregarConfiguracoes() {
  const config = JSON.parse(localStorage.getItem("escalaConfig")) || {
    assistentes: ["Assistente 1", "Assistente 2"],
    mensageiros: {
      A: ["M1", "M2", "M3", "M4"],
      B: ["M5", "M6", "M7", "M8"]
    }
  };

  assistentes = config.assistentes;
  mensageiros = config.mensageiros;

  document.getElementById("assistente1").value = assistentes[0];
  document.getElementById("assistente2").value = assistentes[1];

  ["mA1", "mA2", "mA3", "mA4"].forEach((id, i) => {
    document.getElementById(id).value = mensageiros.A[i];
  });
  ["mB1", "mB2", "mB3", "mB4"].forEach((id, i) => {
    document.getElementById(id).value = mensageiros.B[i];
  });
}

function salvarConfiguracoes() {
  assistentes = [
    document.getElementById("assistente1").value || "Assistente 1",
    document.getElementById("assistente2").value || "Assistente 2"
  ];

  mensageiros.A = [
    document.getElementById("mA1").value || "M1",
    document.getElementById("mA2").value || "M2",
    document.getElementById("mA3").value || "M3",
    document.getElementById("mA4").value || "M4"
  ];

  mensageiros.B = [
    document.getElementById("mB1").value || "M5",
    document.getElementById("mB2").value || "M6",
    document.getElementById("mB3").value || "M7",
    document.getElementById("mB4").value || "M8"
  ];

  const config = { assistantes, mensageiros };
  localStorage.setItem("escalaConfig", JSON.stringify(config));
  carregarMes();
}

function gerarEscalaMensal(mes, ano) {
  escalaCompleta = [];
  const totalDias = new Date(ano, mes + 1, 0).getDate();
  let grupo = "A";

  for (let dia = 1; dia <= totalDias; dia++) {
    const data = new Date(ano, mes, dia);
    const diaSemana = data.toLocaleDateString("pt-BR", { weekday: "long" });
    const grupoId = grupo === "A" ? 0 : 1;

    escalaCompleta.push({
      data: data.toLocaleDateString("pt-BR"),
      dia: diaSemana.charAt(0).toUpperCase() + diaSemana.slice(1),
      assistente: assistentes[grupoId],
      mensageiros: mensageiros[grupo].join(", "),
      grupo
    });

    grupo = grupo === "A" ? "B" : "A";
  }
}

function exibirSemana(semana) {
  const tabela = document.getElementById("tabelaEscala");
  tabela.innerHTML = "";

  const inicio = semana * diasVisiveis;
  const fim = inicio + diasVisiveis;
  const semanaEscala = escalaCompleta.slice(inicio, fim);

  semanaEscala.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td data-label="Data">${item.data}</td>
      <td data-label="Dia">${item.dia}</td>
      <td data-label="Assistente">${item.assistente}</td>
      <td data-label="Mensageiros">${item.mensageiros}</td>
      <td data-label="Grupo" class="grupo-${item.grupo}">Grupo ${item.grupo}</td>
    `;
    tabela.appendChild(tr);
  });
}

function mostrarProximaSemana() {
  const totalSemanas = Math.ceil(escalaCompleta.length / diasVisiveis);
  semanaAtual = (semanaAtual + 1) % totalSemanas;
  exibirSemana(semanaAtual);
}

function mostrarSemanaAnterior() {
  const totalSemanas = Math.ceil(escalaCompleta.length / diasVisiveis);
  semanaAtual = (semanaAtual - 1 + totalSemanas) % totalSemanas;
  exibirSemana(semanaAtual);
}

function carregarMes() {
  const mesSelect = document.getElementById("mesSelect");
  const anoSelect = document.getElementById("anoSelect");

  mesAtual = parseInt(mesSelect.value);
  anoAtual = parseInt(anoSelect.value);
  semanaAtual = 0;

  gerarEscalaMensal(mesAtual, anoAtual);
  exibirSemana(semanaAtual);
}

function exportarExcel() {
  const dados = escalaCompleta.map(item => ({
    Data: item.data,
    Dia: item.dia,
    Assistente: item.assistente,
    Mensageiros: item.mensageiros,
    Grupo: `Grupo ${item.grupo}`
  }));

  const ws = XLSX.utils.json_to_sheet(dados);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Escala");

  const nomeArquivo = `escala_${anoAtual}_${(mesAtual + 1).toString().padStart(2, '0')}.xlsx`;
  XLSX.writeFile(wb, nomeArquivo);
}

// Inicialização
preencherSelects();
carregarConfiguracoes();
carregarMes();
