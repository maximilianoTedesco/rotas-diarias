const inputArquivo = document.getElementById("arquivoExcel");
const btnGerar = document.getElementById("btnGerar");

const resumo = document.getElementById("resumo");
const resultado = document.getElementById("resultado");

const estatisticas = document.getElementById("estatisticas");
const rotasGeradas = document.getElementById("rotasGeradas");

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

        const primeiraPlanilha = workbook.Sheets[
            workbook.SheetNames[0]
        ];

        const linhas = XLSX.utils.sheet_to_json(
            primeiraPlanilha,
            {
                header: 1,
                defval: ""
            }
        );

        processarLinhas(linhas);

    };

    leitor.readAsArrayBuffer(arquivo);

}

function processarLinhas(linhas) {

    const passageiros = [];

    for (let i = 1; i < linhas.length; i++) {

        const linha = linhas[i];

        const entrada = String(
            linha[CONFIG.colunas.entrada]
        ).trim().toUpperCase();

        if (
            entrada === "" ||
            entrada === "FOLGA"
        ) {
            continue;
        }

        passageiros.push({

            nome:
                linha[CONFIG.colunas.nome],

            hotel:
                linha[CONFIG.colunas.hotel],

            entrada:

                linha[CONFIG.colunas.entrada],

            recolha:

                linha[CONFIG.colunas.recolha]

        });

    }

    mostrarResumo(passageiros);

    agruparPorHorario(passageiros);

}

function mostrarResumo(lista) {

    resumo.classList.remove("oculto");

    estatisticas.innerHTML = `

        <div class="estatisticas-grid">

            <div class="stat">

                <strong>${lista.length}</strong>

                Passageiros

            </div>

        </div>

    `;

}

function agruparPorHorario(lista) {

    resultado.classList.remove("oculto");

    rotasGeradas.innerHTML = "";

    const horarios = {};

    lista.forEach(p => {

        if (!horarios[p.entrada]) {

            horarios[p.entrada] = [];

        }

        horarios[p.entrada].push(p);

    });

    Object.keys(horarios)
        .sort()
        .forEach(horario => {

            desenharHorario(

                horario,

                horarios[horario]

            );

        });

}

function desenharHorario(horario, passageiros) {

    const bloco = document.createElement("div");

    bloco.className = "bloco-horario";

    const titulo = document.createElement("div");

    titulo.className = "titulo-horario";

    titulo.innerHTML = `

        Entrada ${horario}

        (${passageiros.length} passageiros)

    `;

    bloco.appendChild(titulo);

    const conteudo = document.createElement("div");

    conteudo.className = "conteudo-horario";

    const hoteis = {};

    passageiros.forEach(p => {

        if (!hoteis[p.hotel]) {

            hoteis[p.hotel] = [];

        }

        hoteis[p.hotel].push(p);

    });

    Object.keys(hoteis)
        .sort()
        .forEach(hotel => {

            const card = document.createElement("div");

            card.className = "hotel";

            card.innerHTML = `

                <h3>

                    🏨 ${hotel}

                    (${hoteis[hotel].length})

                </h3>

            `;

            hoteis[hotel].forEach(p => {

                card.innerHTML += `

                    <div class="passageiro">

                        👤 ${p.nome}

                        <br>

                        📍 ${p.recolha}

                    </div>

                `;

            });

            conteudo.appendChild(card);

        });

    bloco.appendChild(conteudo);

    rotasGeradas.appendChild(bloco);

}
