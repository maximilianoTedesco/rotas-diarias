const inputArquivo = document.getElementById("arquivoExcel");
const btnGerar = document.getElementById("btnGerar");

const resumo = document.getElementById("resumo");
const resultado = document.getElementById("resultado");
const estatisticas = document.getElementById("estatisticas");
const rotasGeradas = document.getElementById("rotasGeradas");

let operacaoFinal = [];

btnGerar.addEventListener("click", lerExcel);

function lerExcel() {
    const arquivo = inputArquivo.files[0];

    if (!arquivo) {
        alert("Selecione um arquivo Excel.");
        return;
    }

    const leitor = new FileReader();

    leitor.onload = function (e) {
        const dados = new Uint8Array(e.target.result);

        const workbook = XLSX.read(dados, {
            type: "array"
        });

        const planilha = workbook.Sheets[workbook.SheetNames[0]];

        const linhas = XLSX.utils.sheet_to_json(planilha, {
            header: 1,
            defval: ""
        });

        processarLinhas(linhas);
    };

    leitor.readAsArrayBuffer(arquivo);
}

function processarLinhas(linhas) {
    const passageiros = [];

    for (let i = 1; i < linhas.length; i++) {
        const linha = linhas[i];

        const nome = String(linha[CONFIG.colunas.nome] || "").trim();
        const hotel = String(linha[CONFIG.colunas.hotel] || "").trim();
        const entrada = normalizarHorario(linha[CONFIG.colunas.entrada]);
        const recolha = String(linha[CONFIG.colunas.recolha] || "").trim();

        if (!nome || !hotel || !entrada) {
            continue;
        }

        if (entrada.toUpperCase().includes("FOLGA")) {
            continue;
        }

        passageiros.push({
            nome,
            hotel,
            entrada,
            recolha
        });
    }

    gerarOperacao(passageiros);
}

function gerarOperacao(passageiros) {
    operacaoFinal = CONFIG.veiculos.map(v => ({
        ...v,
        voltas: []
    }));

    const porHorario = agruparPor(passageiros, "entrada");

    Object.keys(porHorario)
        .sort(ordenarHorariosOperacao)
        .forEach(horario => {

            let fila = [...porHorario[horario]];

            let indiceVeiculo = 0;

            while (fila.length > 0) {
                const veiculosDisponiveis = operacaoFinal.filter(v =>
                    v.voltas.length < v.maxVoltas
                );

                if (veiculosDisponiveis.length === 0) {
                    alert("Não há veículos suficientes para distribuir todos os passageiros.");
                    break;
                }

                const veiculo = veiculosDisponiveis[
                    indiceVeiculo % veiculosDisponiveis.length
                ];

                const passageirosDaVolta = fila.splice(0, veiculo.capacidade);

                veiculo.voltas.push({
                    entrada: horario,
                    saida: calcularSaida(horario),
                    passageiros: passageirosDaVolta
                });

                indiceVeiculo++;
            }
        });

    mostrarResumo(passageiros);
    desenharOperacao();
}
function escolherVeiculoDisponivel() {
    return operacaoFinal.find(v => v.voltas.length < v.maxVoltas);
}

function mostrarResumo(passageiros) {
    resumo.classList.remove("oculto");

    const totalVoltas = operacaoFinal.reduce(
        (total, veiculo) => total + veiculo.voltas.length,
        0
    );

    estatisticas.innerHTML = `
        <div class="estatisticas-grid">
            <div class="stat">
                <strong>${passageiros.length}</strong>
                Passageiros ativos
            </div>

            <div class="stat">
                <strong>${totalVoltas}</strong>
                Voltas geradas
            </div>

            <div class="stat">
                <strong>${CONFIG.veiculos.length}</strong>
                Veículos
            </div>
        </div>
    `;
}

function desenharOperacao() {
    resultado.classList.remove("oculto");
    rotasGeradas.innerHTML = "";

    operacaoFinal.forEach(veiculo => {
        if (veiculo.voltas.length === 0) {
            return;
        }

        const bloco = document.createElement("div");
        bloco.className = "bloco-horario";

        bloco.innerHTML = `
            <div class="titulo-horario">
                🚐 ${veiculo.nome}
                <br>
                ${veiculo.voltas.length} volta(s)
            </div>

            <div class="conteudo-horario">
                ${veiculo.voltas.map((volta, index) => montarHtmlVolta(volta, index)).join("")}

                <button onclick="exportarPdfVeiculo(${veiculo.id})">
                    📄 Exportar PDF deste veículo
                </button>
            </div>
        `;

        rotasGeradas.appendChild(bloco);
    });
}

function montarHtmlVolta(volta, index) {
    const hoteis = agruparPor(volta.passageiros, "hotel");
    const pontos = agruparPor(volta.passageiros, "recolha");

    return `
        <div class="hotel">
            <h3>▶️ ${index + 1}ª VOLTA — ⬆️ ENTRADA ${volta.entrada}</h3>
            <p><strong>Saída prevista:</strong> ${volta.saida}</p>
            <p><strong>Total:</strong> ${volta.passageiros.length} passageiro(s)</p>

            <br>

            <h3>🏨 Hotéis</h3>
            ${Object.keys(hoteis).sort().map(hotel => `
                <div class="passageiro">
                    🏨 ${hotel} — ${hoteis[hotel].length}
                </div>
            `).join("")}

            <br>

            <h3>⏰ Recolhas</h3>
            ${Object.keys(pontos).sort().map((ponto, i) => `
                <div class="passageiro">
                    📍 ${calcularHorarioPonto(volta.saida, i)} — ${ponto || "Sem ponto informado"}
                    <br>
                    ${pontos[ponto].map(p => `• ${p.nome}`).join("<br>")}
                </div>
            `).join("")}
        </div>
    `;
}

function exportarPdfVeiculo(id) {
    const veiculo = operacaoFinal.find(v => v.id === id);

    if (!veiculo) {
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    let y = 15;

    doc.setFontSize(16);
    doc.text(`ROTAS - ${veiculo.nome}`, 10, y);

    y += 10;

    veiculo.voltas.forEach((volta, index) => {
        y = verificarPagina(doc, y);

        doc.setFontSize(13);
        doc.text(`${index + 1}ª VOLTA - ENTRADA ${volta.entrada}`, 10, y);
        y += 7;

        doc.setFontSize(10);
        doc.text(`Saída prevista: ${volta.saida}`, 10, y);
        y += 6;

        doc.text(`Total: ${volta.passageiros.length} passageiros`, 10, y);
        y += 8;

        doc.setFontSize(12);
        doc.text("HOTÉIS", 10, y);
        y += 7;

        const hoteis = agruparPor(volta.passageiros, "hotel");

        Object.keys(hoteis).sort().forEach(hotel => {
            y = verificarPagina(doc, y);
            doc.setFontSize(10);
            doc.text(`${hotel} - ${hoteis[hotel].length}`, 12, y);
            y += 5;
        });

        y += 5;
        y = verificarPagina(doc, y);

        doc.setFontSize(12);
        doc.text("RECOLHAS", 10, y);
        y += 7;

        const pontos = agruparPor(volta.passageiros, "recolha");

        Object.keys(pontos).sort().forEach((ponto, i) => {
            y = verificarPagina(doc, y);

            doc.setFontSize(10);
            doc.text(`${calcularHorarioPonto(volta.saida, i)} - ${ponto || "Sem ponto informado"}`, 12, y);
            y += 5;

            pontos[ponto].forEach(p => {
                y = verificarPagina(doc, y);
                doc.text(`- ${p.nome}`, 16, y);
                y += 5;
            });

            y += 3;
        });

        y += 10;
    });

    doc.save(`rotas-${veiculo.nome.toLowerCase().replaceAll(" ", "-")}.pdf`);
}

function agruparPor(lista, campo) {
    return lista.reduce((grupo, item) => {
        const chave = item[campo] || "Não informado";

        if (!grupo[chave]) {
            grupo[chave] = [];
        }

        grupo[chave].push(item);

        return grupo;
    }, {});
}

function calcularSaida(horario) {
    const [h, m] = horario.split(":").map(Number);

    let total = h * 60 + m - CONFIG.minutosAntesEntrada;

    const [hMin, mMin] = CONFIG.horaMinimaInicio.split(":").map(Number);
    const minimo = hMin * 60 + mMin;

    if (total < minimo) {
        total = minimo;
    }

    return formatarMinutos(total);
}

function calcularHorarioPonto(saida, indice) {
    const [h, m] = saida.split(":").map(Number);

    const total = h * 60 + m + indice * CONFIG.intervaloPontos;

    return formatarMinutos(total);
}

function formatarMinutos(total) {
    const h = String(Math.floor(total / 60)).padStart(2, "0");
    const m = String(total % 60).padStart(2, "0");

    return `${h}:${m}`;
}

function normalizarHorario(valor) {
    if (valor instanceof Date) {
        return valor.toTimeString().slice(0, 5);
    }

    if (typeof valor === "number") {
        const totalMinutos = Math.round(valor * 24 * 60);
        return formatarMinutos(totalMinutos);
    }

    return String(valor || "").trim();
}

function verificarPagina(doc, y) {
    if (y > 280) {
        doc.addPage();
        return 15;
    }

    return y;
}

function ordenarHorariosOperacao(a, b) {
    if (a === "00:00") return 1;
    if (b === "00:00") return -1;

    return converterHorarioParaMinutos(a) - converterHorarioParaMinutos(b);
}

function converterHorarioParaMinutos(horario) {
    const [h, m] = horario.split(":").map(Number);
    return h * 60 + m;
}
