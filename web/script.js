const modos = {
    "Conferência de Nota": [
        { grupo: "COMPRAS", abas: ["COMPRAS SEFAZ", "COMPRAS ALTERDATA"] },
        { grupo: "VENDAS", abas: ["VENDAS SEFAZ", "VENDAS ALTERDATA"] }
    ],
    "Check": [
        { grupo: "COMPRAS", abas: ["COMPRAS SEFAZ", "COMPRAS ALTERDATA", "COMPRAS PRODUTOS"] },
        { grupo: "VENDAS", abas: ["VENDAS SEFAZ", "VENDAS ALTERDATA", "VENDAS PRODUTOS"] }
    ]
};

let modoAtual = "Conferência de Nota";
let arquivosSelecionados = {};

function renderizarGrupos() {
    const container = document.getElementById('gruposContainer');
    container.innerHTML = '';
    arquivosSelecionados = {};
    modos[modoAtual].forEach(grupo => {
        const card = document.createElement('div');
        card.className = "card mb-3";
        const cardBody = document.createElement('div');
        cardBody.className = "card-body";
        cardBody.innerHTML = `<h5 class="card-title">${grupo.grupo}</h5>`;
        grupo.abas.forEach(aba => {
            const row = document.createElement('div');
            row.className = "row align-items-center mb-2";
            row.innerHTML = `
                <div class="col-4">
                    <label class="file-label">${aba}:</label>
                </div>
                <div class="col-5">
                    <input type="file" class="form-control" name="arquivo_${aba}" accept=".xlsx,.csv,.xls" required>
                </div>
            `;
            cardBody.appendChild(row);
        });
        card.appendChild(cardBody);
        container.appendChild(card);
    });
    habilitarBotao();
    // Atualiza labels ao selecionar arquivo
    document.querySelectorAll('input[type="file"]').forEach(input => {
        input.addEventListener('change', function() {
            const aba = this.name.replace('arquivo_', '').replace(/_/g, ' ');
            const label = document.getElementById('label_' + aba.replace(/\s/g, '_'));
            if (this.files.length > 0) {
                label.textContent = this.files[0].name;
                arquivosSelecionados[aba] = this.files[0];
            } else {
                label.textContent = "Nada selecionado";
                delete arquivosSelecionados[aba];
            }
            habilitarBotao();
        });
    });
}

function selecionarModo(modo) {
    modoAtual = modo;
    document.getElementById('btnConferencia').classList.toggle('active', modo === "Conferência de Nota");
    document.getElementById('btnCheck').classList.toggle('active', modo === "Check");
    renderizarGrupos();
    document.getElementById('status').textContent = `Modo selecionado: ${modo}. Aguardando ação...`;
}

function habilitarBotao() {
    // Habilita o botão se todos os arquivos obrigatórios estiverem selecionados
    const totalAbas = modos[modoAtual].reduce((acc, grupo) => acc + grupo.abas.length, 0);
    const btn = document.getElementById('btnAgrupar');
    btn.disabled = Object.keys(arquivosSelecionados).length !== totalAbas;
}

document.getElementById('formAgrupamento').addEventListener('submit', function(e) {
    e.preventDefault();
    // Simula envio e progresso
    document.getElementById('status').textContent = "Processando...";
    const barra = document.getElementById('barraProgresso');
    barra.style.width = "0%";
    barra.textContent = "0%";
    let progresso = 0;
    const interval = setInterval(() => {
        progresso += 10;
        barra.style.width = progresso + "%";
        barra.textContent = progresso + "%";
        if (progresso >= 100) {
            clearInterval(interval);
            document.getElementById('status').textContent = "Concluído! Arquivo agrupado com sucesso.";
            setTimeout(() => {
                barra.style.width = "0%";
                barra.textContent = "0%";
                document.getElementById('status').textContent = "Aguardando ação...";
            }, 3000);
        }
    }, 300);
});

// Inicializa quando o DOM estiver carregado
document.addEventListener('DOMContentLoaded', function() {
    renderizarGrupos();
});