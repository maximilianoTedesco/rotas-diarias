const inputArquivo = document.getElementById("arquivoExcel");
const btnGerar = document.getElementById("btnGerar");

const resumo = document.getElementById("resumo");
const resultado = document.getElementById("resultado");
const estatisticas = document.getElementById("estatisticas");
const rotasGeradas = document.getElementById("rotasGeradas");

let rotasFinais = [];

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
        const workbook = XLSX.read(dados, { type: "array" });
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
        const entrada = String(linha[CONFIG.colunas.entrada] || "").trim();
        const recolha = String(linha[CONFIG.colunas.recolha] || "").trim();

        if (!nome || !hotel || !entrada) continue;

        if (entrada.toUpperCase().includes("FOLGA")) continue;

        passageiros.push({
            nome,
            hotel,
            entrada,
            recolha
        });
    }

    gerarRotas(passageiros);
}

function gerarRotas(passageiros) {
    rotasFinais = [];

    const porHorario = agruparPor(passageiros, "entrada");

    Object.keys(porHorario).sort().forEach(horario => {
        const listaHorario = porHorario[horario];

        let indicePassageiro = 0;

        CONFIG.carrinhas.forEach((carrinha, index) => {
            const passageirosCarrinha = listaHorario.slice(
                indicePassageiro,
                indicePassageiro + carrinha.capacidade
            );

            if (passageirosCarrinha.length > 0) {
                rotasFinais.push({
                    id: rotasFinais.length + 1,
                    horario,
                    saida: calcularSaida(horario),
                    carrinha: carrinha.nome,
                    capacidade: carrinha.capacidade,
                    passageiros: passageirosCarrinha
                });
            }

            indicePassageiro += carrinha.capacidade;
        });

        const restantes = listaHorario.slice(indicePassageiro);

        if (restantes.length > 0) {
            rotasFinais.push({
                id: rotasFinais.length + 1,
                horario,
                saida: calcularSaida(horario),
                carrinha: "Viagem extra necessária",
                capacidade: restantes.length,
                passageiros: restantes,
                extra: true
            });
        }
    });

    mostrarResumo(passageiros);
    desenharRotas();
}

function mostrarResumo(passageiros) {
    resumo.classList.remove("oculto");

    estatisticas.innerHTML = `
        <div class="estatisticas-grid">
            <div class="stat">
                <strong>${passageiros.length}</strong>
                Passageiros ativos
            </div>

            <div class="stat">
                <strong>${rotasFinais.length}</strong>
                Rotas geradas
            </div>

            <div class="stat">
                <strong>${CONFIG.carrinhas.length}</strong>
                Carrinhas disponíveis
            </div>
        </div>
    `;
}

function desenharRotas() {
    resultado.classList.remove("oculto");
    rotasGeradas.innerHTML = "";

    rotasFinais.forEach(rota => {
        const bloco = document.createElement("div");
        bloco.className = "bloco-horario";

        const hoteis = agruparPor(rota.passageiros, "hotel");
        const pontos = agruparPor(rota.passageiros, "recolha");

        bloco.innerHTML = `
            <div class="titulo-horario">
                Rota ${rota.id} — Entrada ${rota.horario}
                <br>
                Saída prevista: ${rota.saida}
            </div>

            <div class="conteudo-horario">
                <p><strong>Carrinha:</strong> ${rota.carrinha}</p>
                <p><strong>Capacidade:</strong> ${rota.capacidade}</p>
                <p><strong>Total:</strong> ${rota.passageiros.length} passageiros</p>

                ${rota.extra ? `<div class="alerta">Atenção: esta rota excede a capacidade normal da frota.</div>` : ""}

                <br>

                <h3>🏨 Hotéis</h3>
                ${Object.keys(hoteis).sort().map(hotel => `
                    <div class="hotel">
                        <h3>${hotel} (${hoteis[hotel].length})</h3>
                        ${hoteis[hotel].map(p => `
                            <div class="passageiro">
                                👤 ${p.nome}<br>
                                📍 ${p.recolha}
                            </div>
                        `).join("")}
                    </div>
                `).join("")}

                <h3>📍 Pontos de recolha</h3>
                ${Object.keys(pontos).sort().map(ponto => `
                    <div class="hotel">
                        <h3>${ponto || "Sem ponto informado"} (${pontos[ponto].length})</h3>
                        ${pontos[ponto].map(p => `
                            <div class="passageiro">
                                👤 ${p.nome}<br>
                                🏨 ${p.hotel}
                            </div>
                        `).join("")}
                    </div>
                `).join("")}

                <button onclick="exportarPdf(${rota.id})">
                    📄 Exportar PDF desta rota
                </button>
            </div>
        `;

        rotasGeradas.appendChild(bloco);
    });
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
    const partes = String(horario).split(":");

    let horas = Number(partes[0]);
    let minutos = Number(partes[1] || 0);

    let totalMinutos = horas * 60 + minutos;
    totalMinutos -= CONFIG.minutosAntesEntrada;

    if (totalMinutos < 0) {
        totalMinutos += 24 * 60;
    }

    const h = String(Math.floor(totalMinutos / 60)).padStart(2, "0");
    const m = String(totalMinutos % 60).padStart(2, "0");

    return `${h}:${m}`;
}

function exportarPdf(id) {
    const rota = rotasFinais.find(r => r.id === id);

    if (!rota) return;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    let y = 15;

    doc.setFontSize(16);
    doc.text(`ROTA ${rota.id} - ENTRADA ${rota.horario}`, 10, y);

    y += 10;

    doc.setFontSize(11);
    doc.text(`Saída prevista: ${rota.saida}`, 10, y);
    y += 7;
    doc.text(`Carrinha: ${rota.carrinha}`, 10, y);
    y += 7;
    doc.text(`Capacidade: ${rota.capacidade}`, 10, y);
    y += 7;
    doc.text(`Total: ${rota.passageiros.length} passageiros`, 10, y);

    y += 12;

    doc.setFontSize(13);
    doc.text("HOTÉIS", 10, y);
    y += 8;

    const hoteis = agruparPor(rota.passageiros, "hotel");

    Object.keys(hoteis).sort().forEach(hotel => {
        y = verificarPagina(doc, y);

        doc.setFontSize(11);
        doc.text(`${hotel} (${hoteis[hotel].length})`, 10, y);
        y += 6;

        hoteis[hotel].forEach(p => {
            y = verificarPagina(doc, y);
            doc.setFontSize(10);
            doc.text(`- ${p.nome} | ${p.recolha}`, 14, y);
            y += 5;
        });

        y += 4;
    });

    y += 6;
    y = verificarPagina(doc, y);

    doc.setFontSize(13);
    doc.text("PONTOS DE RECOLHA", 10, y);
    y += 8;

    const pontos = agruparPor(rota.passageiros, "recolha");

    Object.keys(pontos).sort().forEach(ponto => {
        y = verificarPagina(doc, y);

        doc.setFontSize(11);
        doc.text(`${ponto || "Sem ponto informado"} (${pontos[ponto].length})`, 10, y);
        y += 6;

        pontos[ponto].forEach(p => {
            y = verificarPagina(doc, y);
            doc.setFontSize(10);
            doc.text(`- ${p.nome} | ${p.hotel}`, 14, y);
            y += 5;
        });

        y += 4;
    });

    doc.save(`rota-${rota.id}-entrada-${rota.horario}.pdf`);
}

function verificarPagina(doc, y) {
    if (y > 280) {
        doc.addPage();
        return 15;
    }

    return y;
}
