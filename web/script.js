let currentMode = 'conferencia';
let selectedFiles = {};
let selectedDirectory = null;
let isProcessing = false;
    
const modes = {
    'conferencia': [
        {
            name: 'COMPRAS',
            tabs: ['COMPRAS SEFAZ', 'COMPRAS ALTERDATA']
        },
        {
            name: 'VENDAS',
            tabs: ['VENDAS SEFAZ', 'VENDAS ALTERDATA']
        }
    ],
    'check': [
        {
            name: 'COMPRAS',
            tabs: ['COMPRAS SEFAZ', 'COMPRAS ALTERDATA', 'COMPRAS PRODUTOS']
        },
        {
            name: 'VENDAS',
            tabs: ['VENDAS SEFAZ', 'VENDAS ALTERDATA', 'VENDAS PRODUTOS']
        }
    ]
};

function selectMode(mode) {
    if (isProcessing) {
        alert('Processo em andamento. Aguarde terminar para mudar o modo.');
        return;
    }
    
    currentMode = mode;
    selectedFiles = {};
    selectedDirectory = null;
    
    // Atualizar botões de modo
    document.getElementById('conferencia-btn').classList.remove('active');
    document.getElementById('check-btn').classList.remove('active');
    document.getElementById(mode + '-btn').classList.add('active');
    
    // Limpar diretório
    document.getElementById('directory-path').textContent = 'Nenhum diretório selecionado';
    document.getElementById('confirm-btn').disabled = true;
    
    // Resetar nome do arquivo
    document.getElementById('filename-input').value = 'planilhas_agrupadas';
    
    // Atualizar status
    const modeText = mode === 'conferencia' ? 'Conferência de Nota' : 'Check';
    document.getElementById('status-label').textContent = `Modo selecionado: ${modeText}. Aguardando ação...`;
    
    // Criar seções de arquivos
    createFileSections();
}

function createFileSections() {
    const container = document.getElementById('file-sections');
    container.innerHTML = '';
    
    const groups = modes[currentMode];
    
    groups.forEach(group => {
        const section = document.createElement('div');
        section.className = 'group-section';
        
        const title = document.createElement('div');
        title.className = 'group-title';
        title.textContent = group.name;
        section.appendChild(title);
        
        group.tabs.forEach(tab => {
            const row = document.createElement('div');
            row.className = 'file-row';
            
            const btn = document.createElement('button');
            btn.className = 'file-btn';
            btn.textContent = tab;
            btn.onclick = () => selectFile(tab);
            
            const status = document.createElement('div');
            status.className = 'file-status';
            status.id = 'status-' + tab.replace(/\s+/g, '-');
            status.textContent = 'Nada selecionado';
            
            row.appendChild(btn);
            row.appendChild(status);
            section.appendChild(row);
        });
        
        container.appendChild(section);
    });
}

function selectFile(tabName) {
    // Simular seleção de arquivo (em uma aplicação real, usaria input file)
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.csv,.xls';
    input.onchange = function(e) {
        if (e.target.files.length > 0) {
            const file = e.target.files[0];
            selectedFiles[tabName] = file;
            const statusId = 'status-' + tabName.replace(/\s+/g, '-');
            document.getElementById(statusId).textContent = file.name;
        }
    };
    input.click();
}

function selectDirectory() {
    
    // Simular seleção de diretório (em uma aplicação real, usaria API específica)
    const path = prompt('Digite o caminho do diretório de destino:');
    if (path) {
        selectedDirectory = path;
        document.getElementById('directory-path').textContent = path;
        document.getElementById('confirm-btn').disabled = false;
    }
}
    
async function confirmProcess() {
    try {
        const formData = new FormData();
        for (const [tabName, file] of Object.entries(selectedFiles)) {
        formData.append('exceis', file);
        formData.append('aba', tabName);
        }
        formData.append('nome_saida', document.getElementById('filename-input').value || 'planilhas_agrupadas');
        formData.append('modo', currentMode)

        await fetch('http://127.0.0.1:8000/processar', { method: 'POST', body: formData });
    } catch (error) {
        console.log('Network error:', error);
    }
}
    
function simulateProcess() {
    const steps = [
        'Iniciando agrupamento dos arquivos...',
        'Agrupando arquivos',
        'Processando SEFAZ',
        'Processando ALTERDATA',
        'Processando PRODUTO',
        'Verificando COMPRAS',
        'Verificando VENDAS'
    ];
    
    let currentStep = 0;
    const progressFill = document.getElementById('progress-fill');
    const statusLabel = document.getElementById('status-label');
    
    const interval = setInterval(() => {
        if (currentStep < steps.length) {
            statusLabel.textContent = steps[currentStep] + ' - Tempo decorrido: ' + (currentStep * 2 + 2) + '.0 s';
            progressFill.style.width = ((currentStep + 1) / steps.length * 100) + '%';
            currentStep++;
        } else {
            clearInterval(interval);
            
            // Finalizar processo
            statusLabel.textContent = 'Concluído em 14.00 segundos. Processo finalizado.';
            document.getElementById('confirm-btn').disabled = false;
            document.getElementById('confirm-btn').textContent = 'Agrupar em um Excel';
            progressFill.style.width = '0%';
            isProcessing = false;
            
            // Limpar seleções
            selectedFiles = {};
            selectedDirectory = null;
            document.getElementById('directory-path').textContent = 'Nenhum diretório selecionado';
            document.getElementById('filename-input').value = 'planilhas_agrupadas';
            document.getElementById('confirm-btn').disabled = true;
            
            // Resetar status dos arquivos
            const statusElements = document.querySelectorAll('.file-status');
            statusElements.forEach(el => {
                el.textContent = 'Nada selecionado';
            });
        }
    }, 2000);
}
    
// Inicializar interface
createFileSections();