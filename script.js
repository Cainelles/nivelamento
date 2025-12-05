let chartAusetu, chartCaldtu, chartMandtu, chartTotal;

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

// Carrega a planilha automaticamente ao abrir a página
window.addEventListener('load', carregarPlanilhaAutomatico);

// Atualiza os gráficos quando o tipo de gráfico muda
document.getElementById("tipoGrafico").addEventListener("change", carregarPlanilhaAutomatico);

function carregarPlanilhaAutomatico() {
    let tipoGrafico = document.getElementById("tipoGrafico").value || "bar";

    fetch('planilha.xlsx')
        .then(res => res.arrayBuffer())
        .then(data => {
            let workbook = XLSX.read(data, { type: "array" });
            let sheet = workbook.Sheets[workbook.SheetNames[0]];

            function getPercent(cell) {
                let val = sheet[cell]?.v || 0;
                val = typeof val === "number" ? val : parseFloat(val.toString().replace('%','').replace(',','.'));
                return val * 100;
            }

            let AUSETU = {
                labels: ["TRANSF. ESTOC.", "UTILIDADES"],
                valores: [getPercent("K20"), getPercent("K21")]
            };

            let CALDTU = {
                labels: ["TRANSF. ESTOC.", "UTILIDADES"],
                valores: [getPercent("K31"), getPercent("K32")]
            };

            let MANDTU = {
                labels: ["TRANSF. ESTOC.", "UTILIDADES"],
                valores: [getPercent("K42"), getPercent("K43")]
            };

            let TOTAL = {
                labels: ["TRANSF. ESTOC.", "UTILIDADES"],
                valores: [getPercent("K51"), getPercent("K52")]
            };

            gerarGrafico("chartAusetu", AUSETU, chartAusetu, (c) => chartAusetu = c, tipoGrafico);
            gerarGrafico("chartCaldtu", CALDTU, chartCaldtu, (c) => chartCaldtu = c, tipoGrafico);
            gerarGrafico("chartMandtu", MANDTU, chartMandtu, (c) => chartMandtu = c, tipoGrafico);
            gerarGrafico("chartTotal", TOTAL, chartTotal, (c) => chartTotal = c, tipoGrafico);
        })
        .catch(err => console.error("Erro ao carregar a planilha:", err));
}

function gerarGrafico(canvasId, dados, oldChart, saveChart, tipoGrafico) {
    if (oldChart) oldChart.destroy();

    const isEscuro = temaSelect.value === 'escuro';

    const coresClaro = ["rgba(0, 206, 209, 0.6)", "rgba(135, 102, 255, 0.6)"];
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
            maintainAspectRatio: false,
            plugins: {
                legend: { 
                    display: true,
                    labels: { color: datalabelColor }
                },
                tooltip: { enabled: true },
                datalabels: {
                    anchor: tipoGrafico === "pie" ? 'center' : 'end',
                    align: tipoGrafico === "pie" ? 'center' : 'end',
                    formatter: (value) => Math.round(value) + '%',
                    color: datalabelColor,
                    font: { 
                        weight: 'bold', 
                        size: tipoGrafico === "pie" ? 20 : 14
                    }
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

// Atualiza gráficos existentes ao trocar de tema
function atualizarTemaGrafico() {
    const charts = [chartAusetu, chartCaldtu, chartMandtu, chartTotal];
    charts.forEach(chart => {
        if (!chart) return;

        const isEscuro = temaSelect.value === 'escuro';

        const coresClaro = ["rgba(0, 206, 209, 0.6)", "rgba(135, 102, 255, 0.6)"];
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
