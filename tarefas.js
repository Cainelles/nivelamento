let dadosAba = [];

const btnIndex = document.getElementById("btnIndex");

// Seleção de tema
const temaSelect = document.getElementById('tema');

temaSelect.addEventListener('change', () => {
    if (temaSelect.value === 'escuro') {
        document.body.classList.add('escuro');
    } else {
        document.body.classList.remove('escuro');
    }

    // Atualiza os gráficos para o novo tema
    atualizarTemaGrafico();
});

// Elementos
const temaSel = document.getElementById("tema");
const abaSel = document.getElementById("abaPlanilha");
const filtroOrdem = document.getElementById("filtroOrdem");
const filtroTexto = document.getElementById("filtroTexto");
const tabela = document.getElementById("tabelaDados");
const btnFiltrar = document.getElementById("btnFiltrar");

// Eventos
window.addEventListener("load", carregarAba);
abaSel.addEventListener("change", carregarAba);
btnFiltrar.addEventListener("click", aplicarFiltros);
tipoSel.addEventListener("change", aplicarFiltros);
temaSel.addEventListener("change", alternarTema);

// Lê aba da planilha
function carregarAba() {
    fetch("planilha.xlsx")
        .then(r => r.arrayBuffer())
        .then(buffer => {
            const wb = XLSX.read(buffer, { type: "array" });
            const sheet = wb.Sheets[abaSel.value];

            if (!sheet) return alert("Aba não encontrada!");

            dadosAba = [];
            const range = XLSX.utils.decode_range(sheet["!ref"]);

            for (let row = 2; row <= range.e.r; row++) {
                dadosAba.push({
                    conf: get(sheet, "A", row),
                    ordem: get(sheet, "B", row),
                    op: get(sheet, "C", row),
                    subop: get(sheet, "D", row),
                    centro: get(sheet, "E", row),
                    texto: get(sheet, "F", row),
                    exec: get(sheet, "G", row),
                    trab: get(sheet, "H", row) || "",
                    trabReal: get(sheet, "I", row) || "",
                    duracao: get(sheet, "J", row)
                });
            }

            preencherTabela(dadosAba);
           
        });
}

function get(sheet, col, row) {
    const ref = col + (row + 1);
    return sheet[ref] ? sheet[ref].v : "";
}

// Tabela
function preencherTabela(lista) {
    tabela.innerHTML = "";
    lista.forEach(l => {
        tabela.innerHTML += `
            <tr>
                <td>${l.conf}</td>
                <td>${l.ordem}</td>
                <td>${l.op}</td>
                <td>${l.subop}</td>
                <td>${l.centro}</td>
                <td>${l.texto}</td>
                <td>${l.exec}</td>
                <td>${l.trab}</td>
                <td>${l.trabReal}</td>
                <td>${l.duracao}</td>
            </tr>
        `;
    });
}

// Filtros
function aplicarFiltros() {
    let filtrado = dadosAba;

    if (filtroOrdem.value.trim()) {
        filtrado = filtrado.filter(l =>
            String(l.ordem).includes(filtroOrdem.value.trim())
        );
    }

    if (filtroTexto.value.trim()) {
        filtrado = filtrado.filter(l =>
            String(l.texto).toLowerCase()
                .includes(filtroTexto.value.trim().toLowerCase())
        );
    }

    preencherTabela(filtrado);
 
}



// Tema
function alternarTema() {
    document.body.classList.toggle("escuro", temaSel.value === "escuro");
    aplicarFiltros();
}


