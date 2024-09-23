document.getElementById('fileInput').addEventListener('change', processarArquivo);

function processarArquivo() {
    const input = document.getElementById('fileInput');
    const reader = new FileReader();

    reader.onload = function (event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const baseQiveSheet = workbook.Sheets['Base Qive'];
            const loteSheet = workbook.Sheets['Lote'];
            if (!baseQiveSheet) throw new Error("Planilha 'Base Qive' não encontrada.");
            if (!loteSheet) throw new Error("Planilha 'Lote' não encontrada.");

            const baseQiveData = XLSX.utils.sheet_to_json(baseQiveSheet, { raw: true });
            const loteData = XLSX.utils.sheet_to_json(loteSheet, { raw: true });

            const parseData = (valor) => {
                if (typeof valor === 'number') {
                    return formatarData(serialParaData(valor));
                }

                if (typeof valor === 'string') {
                    const partes = valor.split('/');
                    if (partes.length === 3) {
                        const date = new Date(partes[2], partes[1] - 1, partes[0]);
                        return formatarData(date);
                    }
                }
                return null;
            };

            const serialParaData = (serial) => {
                const baseDate = new Date(Date.UTC(1899, 11, 30));
                return new Date(baseDate.getTime() + serial * 86400000);
            };

            const formatarData = (date) => {
                const dia = String(date.getDate()).padStart(2, '0');
                const mes = String(date.getMonth() + 1).padStart(2, '0');
                const ano = date.getFullYear();
                return `${dia}/${mes}/${ano}`;
            };

            // Mapeando os dados da Base Qive
            const baseQiveMap = baseQiveData.reduce((acc, item) => {
                const chave = item.CNPJ + '-' + (item['Valor'] || '');
                acc[chave] = { codigo: item.Código, competencia: parseData(item.Competência), valor: item['Valor'] };
                return acc;
            }, {});

            const duplicidadeCount = baseQiveData.reduce((acc, item) => {
                const chave = item.CNPJ + '-' + (item['Valor'] || '');
                acc[chave] = (acc[chave] || 0) + 1;
                return acc;
            }, {});

            const totalValoresPorCNPJ = baseQiveData.reduce((acc, item) => {
                const cnpj = item.CNPJ;
                const valor = parseFloat(item['Valor']) || 0;
                acc[cnpj] = (acc[cnpj] || 0) + valor;
                return acc;
            }, {});

            const notasCount = baseQiveData.reduce((acc, item) => {
                const cnpj = item.CNPJ;
                acc[cnpj] = (acc[cnpj] || 0) + 1;
                return acc;
            }, {});

            const tabela = document.getElementById('table-lote').querySelector('tbody');
            tabela.innerHTML = '';

            loteData.forEach(lote => {
                const tr = document.createElement('tr');

                // Adicionando as colunas com os dados
                tr.innerHTML += `
                    <td>${lote.Nome || ''}</td>
                    <td>${lote.CNPJ || ''}</td>
                    <td>${lote['NFS-e'] || ''}</td>
                `;

                const chave = lote.CNPJ + '-' + (lote['NFS-e'] || '');
                const dadosCorrespondentes = baseQiveMap[chave];
                const fechamentoTd = document.createElement('td');
                const codigoTd = document.createElement('td');
                const competenciaTd = document.createElement('td');
                const duplicidadeTd = document.createElement('td');
                const valorTotalTd = document.createElement('td');
                const notasTd = document.createElement('td');

                const valorTotal = totalValoresPorCNPJ[lote.CNPJ] || 0;
                valorTotalTd.textContent = valorTotal.toFixed(2);

                if (dadosCorrespondentes) {
                    const { codigo, competencia } = dadosCorrespondentes;
                    const quantidade = duplicidadeCount[chave] || 1;

                    let status = quantidade > 1 ? 'Não Apto' : 'Apto';

                    codigoTd.textContent = codigo || '';
                    competenciaTd.textContent = competencia || '';
                    duplicidadeTd.textContent = quantidade;
                    fechamentoTd.textContent = status;

                    tr.classList.add(status === 'Apto' ? 'apto' : 'nao-apto');
                } else {
                    codigoTd.textContent = '';
                    competenciaTd.textContent = '';
                    duplicidadeTd.textContent = '1';
                    fechamentoTd.textContent = 'Não Apto';
                    tr.classList.add('nao-apto');
                }

                notasTd.textContent = notasCount[lote.CNPJ] || 0; // Adicionando a contagem de notas

                tr.appendChild(valorTotalTd);
                tr.appendChild(notasTd); // Mover a nova coluna antes da coluna de fechamento
                tr.appendChild(codigoTd);
                tr.appendChild(competenciaTd);
                tr.appendChild(duplicidadeTd);
                tr.appendChild(fechamentoTd); // Fechamento como a última coluna
                tabela.appendChild(tr);
            });
        } catch (error) {
            console.error("Erro ao processar o arquivo:", error);
        }
    };

    reader.readAsArrayBuffer(input.files[0]);
}

function filtrarTabela() {
    const filtro = document.getElementById('statusFilter').value;
    const linhas = document.querySelectorAll('#table-lote tbody tr');

    linhas.forEach(linha => {
        const status = linha.lastChild.textContent.trim(); 
        linha.style.display = (filtro === '' || status === filtro) ? '' : 'none';
    });
}

function baixarAptos() {
    const aptos = [];
    const linhas = document.querySelectorAll('#table-lote tbody tr');

    linhas.forEach(linha => {
        if (linha.classList.contains('apto')) {
            const dados = Array.from(linha.children).map(td => td.textContent);
            aptos.push(dados);
        }
    });

    if (aptos.length > 0) {
        const ws = XLSX.utils.aoa_to_sheet([['Nome', 'CNPJ', 'NFS-e', 'Valor Total', 'Notas', 'Código', 'Competência', 'Duplicidade', 'Status'], ...aptos]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Entregadores Aptos');
        XLSX.writeFile(wb, 'entregadores_aptos.xlsx');
    } else {
        alert('Nenhum entregador apto encontrado.');
    }
}

document.getElementById('baixarPlanilhaCompleta').addEventListener('click', baixarPlanilhaCompleta);

function baixarPlanilhaCompleta() {
    const headers = ['Nome', 'CNPJ', 'NFS-e', 'Valor Total', 'Notas', 'Código', 'Competência', 'Duplicidade', 'Fechamento'];
    const dadosTabela = [];
    const linhas = document.querySelectorAll('#table-lote tbody tr');

    linhas.forEach(linha => {
        const dados = Array.from(linha.children).map(td => td.textContent);
        dadosTabela.push(dados);
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...dadosTabela]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Planilha Completa');
    XLSX.writeFile(wb, 'planilha_completa.xlsx');
}
