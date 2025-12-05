let chartAusetu, chartCaldtu, chartMandtu, chartTotal;

const temaSelect = document.getElementById('tema');
const tipoGraficoSelect = document.getElementById('tipoGrafico');

// Troca de tema
temaSelect.addEventListener('change', () => {
    document.body.classList.toggle("escuro", temaSelect.value === 'escuro');
    atualizarTemaGrafico();
});

// Troca de tipo de gráfico
tipoGraficoSelect.addEventListener("change", carregarPlanilhaAutomatico);

// Carrega planilha e gráficos ao abrir
window.addEventListener('load', carregarPlanilhaAutomatico);

function carregarPlanilhaAutomatico() {
    const tipoGrafico = tipoGraficoSelect.value || "bar";

    fetch('planilha.xlsx')
        .then(res => res.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: "array" });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            // Função para pegar valor e ignorar células vazias
            function getPercent(cell) {
                const val = sheet[cell]?.v;
                if (val === undefined || val === null || val === "") return null; // ignora vazio
                let num = typeof val === "number" ? val : parseFloat(val.toString().replace('%','').replace(',','.')) || 0;
                return num * 100;
            }

            function prepararDados(cells) {
                const valores = cells.map(getPercent).filter(v => v !== null);
                const labels = valores.length === 2 ? ["TRANSF. ESTOC.", "UTILIDADES"] : [];
                return { labels, valores };
            }

            const AUSETU = prepararDados(["K20", "K21"]);
            const CALDTU = prepararDados(["K31", "K32"]);
            const MANDTU = prepararDados(["K42", "K43"]);
            const TOTAL  = prepararDados(["K51", "K52"]);

            gerarGrafico("chartAusetu", AUSETU, chartAusetu, c => chartAusetu = c, tipoGrafico);
            gerarGrafico("chartCaldtu", CALDTU, chartCaldtu, c => chartCaldtu = c, tipoGrafico);
            gerarGrafico("chartMandtu", MANDTU, chartMandtu, c => chartMandtu = c, tipoGrafico);
            gerarGrafico("chartTotal", TOTAL, chartTotal, c => chartTotal = c, tipoGrafico);
        })
        .catch(err => console.error("Erro ao carregar a planilha:", err));
}

function gerarGrafico(canvasId, dados, oldChart, saveChart, tipoGrafico) {
    if (oldChart) oldChart.destroy();

    const isEscuro = temaSelect.value === 'escuro';
    const coresClaro = ["rgba(0, 206, 209, 0.6)", "rgba(255, 255, 0, 0.6)"];
    const coresEscuro = ["rgba(238, 0, 0, 0.8)", "rgba(145, 44, 238, 0.8)"];
    const backgroundColors = isEscuro ? coresEscuro : coresClaro;
    const borderColor = isEscuro ? "#ffffff" : "rgba(0,0,0,0.9)";
    const datalabelColor = isEscuro ? "#ffffff" : "#000000";

    const config = {
        type: tipoGrafico,
        data: {
            labels: dados.labels,
            datasets: [{
                label: "Porcentagens",
                data: dados.valores,
                backgroundColor: backgroundColors,
                borderColor: borderColor,
                borderWidth: 1,
                fill: tipoGrafico === "line" ? false : true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            aspectRatio: 1, // mantém proporção quadrada para pizza/linha
            plugins: {
                legend: { display: true, labels: { color: datalabelColor } },
                tooltip: { enabled: true },
                datalabels: {
                    anchor: tipoGrafico === "pie" ? 'center' : 'end',
                    align: tipoGrafico === "pie" ? 'center' : 'end',
                    formatter: (value) => Math.round(value) + '%',
                    color: datalabelColor,
                    font: { weight: 'bold', size: tipoGrafico === "pie" ? 20 : 14 }
                }
            },
            scales: tipoGrafico === "pie" ? {} : { 
                y: { beginAtZero: true, ticks: { color: datalabelColor } },
                x: { ticks: { color: datalabelColor } }
            }
        },
        plugins: [ChartDataLabels]
    };

    saveChart(new Chart(document.getElementById(canvasId), config));
}

function atualizarTemaGrafico() {
    const charts = [chartAusetu, chartCaldtu, chartMandtu, chartTotal];
    charts.forEach(chart => {
        if (!chart) return;

        const isEscuro = temaSelect.value === 'escuro';
        const coresClaro = ["rgba(0, 206, 209, 0.6)", "rgba(255, 255, 0, 0.6)"];
        const coresEscuro = ["rgba(238, 0, 0, 0.8)", "rgba(145, 44, 238, 0.8)"];
        const backgroundColors = isEscuro ? coresEscuro : coresClaro;
        const borderColor = isEscuro ? "#ffffff" : "rgba(0,0,0,0.9)";
        const datalabelColor = isEscuro ? "#ffffff" : "#000000";

        chart.data.datasets[0].backgroundColor = backgroundColors;
        chart.data.datasets[0].borderColor = borderColor;
        chart.options.plugins.datalabels.color = datalabelColor;
        chart.options.plugins.legend.labels.color = datalabelColor;

        if (chart.options.scales?.y) chart.options.scales.y.ticks.color = datalabelColor;
        if (chart.options.scales?.x) chart.options.scales.x.ticks.color = datalabelColor;

        chart.update();
    });
}
