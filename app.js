// ═══════════════════════════════════════════════════════════════════════════
//                              VARIABLES GLOBALES
// ═══════════════════════════════════════════════════════════════════════════

let cashChart = null;
const fileCache = new Map();

// ═══════════════════════════════════════════════════════════════════════════
//                              LOADER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

const LOADER_STEPS = [
    { id: 'magnitude', label: 'Lecture fichiers Magnitude' },
    { id: 'kyriba', label: 'Lecture fichiers Kyriba' },
    { id: 'process', label: 'Traitement des données' },
    { id: 'display', label: 'Affichage' }
];

function showLoader(message = 'Chargement...') {
    let loader = document.getElementById('loader-overlay');
    
    if (!loader) {
        loader = document.createElement('div');
        loader.id = 'loader-overlay';
        loader.innerHTML = `
            <div class="loader-content">
                <div class="loader-logo">🦐</div>
                <p id="loader-message">${message}</p>
                <div class="progress-bar-container">
                    <div class="progress-bar" id="progress-bar"></div>
                </div>
                <div class="loader-steps" id="loader-steps">
                    ${LOADER_STEPS.map(step => `
                        <div class="loader-step" data-step="${step.id}">
                            <span class="loader-step-icon">○</span>
                            <span>${step.label}</span>
                        </div>
                    `).join('')}
                </div>
            </div>
        `;
        document.body.appendChild(loader);
    } else {
        loader.style.display = 'flex';
        document.getElementById('loader-message').textContent = message;
        // Reset steps
        document.querySelectorAll('.loader-step').forEach(step => {
            step.classList.remove('active', 'done');
            step.querySelector('.loader-step-icon').textContent = '○';
        });
    }
}

function setLoaderStep(stepId) {
    const steps = document.querySelectorAll('.loader-step');
    let foundCurrent = false;
    
    steps.forEach(step => {
        const id = step.dataset.step;
        const icon = step.querySelector('.loader-step-icon');
        
        if (id === stepId) {
            step.classList.remove('done');
            step.classList.add('active');
            icon.textContent = '●';
            foundCurrent = true;
        } else if (!foundCurrent) {
            step.classList.remove('active');
            step.classList.add('done');
            icon.textContent = '✓';
        } else {
            step.classList.remove('active', 'done');
            icon.textContent = '○';
        }
    });
    
    const currentStep = LOADER_STEPS.find(s => s.id === stepId);
    if (currentStep) {
        document.getElementById('loader-message').textContent = currentStep.label + '...';
    }
}

function hideLoader() {
    const loader = document.getElementById('loader-overlay');
    if (loader) {
        document.querySelectorAll('.loader-step').forEach(step => {
            step.classList.remove('active');
            step.classList.add('done');
            step.querySelector('.loader-step-icon').textContent = '✓';
        });
        document.getElementById('loader-message').textContent = 'Terminé !';
        
        setTimeout(() => {
            loader.style.display = 'none';
        }, 400);
    }
}

// ═══════════════════════════════════════════════════════════════════════════
//                              INITIALISATION
// ═══════════════════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('folder-input').addEventListener('change', handleFolderSelect);
    document.getElementById('year-select').addEventListener('change', handleYearSelect);
    document.getElementById('month-select').addEventListener('change', handleMonthSelect);
    document.getElementById('entity-select').addEventListener('change', handleEntitySelect);
});

// ═══════════════════════════════════════════════════════════════════════════
//                         CACHE DES FICHIERS EXCEL
// ═══════════════════════════════════════════════════════════════════════════

async function parseExcelFile(file) {
    const cacheKey = file.name + '_' + file.lastModified;
    
    if (fileCache.has(cacheKey)) {
        return fileCache.get(cacheKey);
    }
    
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            fileCache.set(cacheKey, workbook);
            resolve(workbook);
        };
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(file);
    });
}

// ═══════════════════════════════════════════════════════════════════════════
//                         PARTIE SELECTION FICHIER
// ═══════════════════════════════════════════════════════════════════════════

async function handleFolderSelect(event) {
    const files = event.target.files;
    if (files.length === 0) return;
    
    showLoader('Analyse des fichiers...');
    await new Promise(resolve => setTimeout(resolve, 50));
    
    const years = new Set();

    for (const file of files) {
        if (file.webkitRelativePath.includes('/Magnitude/')) {
            const match = file.name.match(/^(\d{4})_\d{2}\.(xlsx)$/);
            if (match) {
                years.add(match[1]);
            }
        }
    }

    const yearSelect = document.getElementById('year-select');
    yearSelect.innerHTML = '<option value="">--Choisir une année--</option>';
    
    const sortedYears = Array.from(years).sort((a, b) => b - a);
    sortedYears.forEach(year => {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        yearSelect.appendChild(option);
    });

    document.getElementById('month-select').innerHTML = '<option value="">--Choisir un mois--</option>';
    document.getElementById('entity-select').innerHTML = '<option value="">--Choisir une entité--</option>';
    
    hideLoader();
}

function handleYearSelect() {
    const selectedYear = document.getElementById('year-select').value;
    if (!selectedYear) return;
    
    const files = document.getElementById('folder-input').files;
    const months = new Set();

    for (const file of files) {
        if (file.webkitRelativePath.includes('/Magnitude/') && file.name.startsWith(selectedYear)) {
            const match = file.name.match(/^\d{4}_(\d{2})\.(xlsx)$/);
            if (match) {
                months.add(match[1]);
            }
        }
    }

    const monthSelect = document.getElementById('month-select');
    monthSelect.innerHTML = '<option value="">--Choisir un mois--</option>';
    
    const sortedMonths = Array.from(months).sort((a, b) => a - b);
    sortedMonths.forEach(month => {
        const option = document.createElement('option');
        option.value = month;
        option.textContent = new Date(selectedYear, month - 1).toLocaleString('fr-FR', { month: 'long' });
        monthSelect.appendChild(option);
    });

    document.getElementById('entity-select').innerHTML = '<option value="">--Choisir une entité--</option>';
}

async function handleMonthSelect() {
    const selectedYear = document.getElementById('year-select').value;
    const selectedMonth = document.getElementById('month-select').value;
    if (!selectedYear || !selectedMonth) return;
    
    showLoader('Chargement des entités...');
    await new Promise(resolve => setTimeout(resolve, 50));
    
    const files = document.getElementById('folder-input').files;
    let selectedFile = null;

    for (const file of files) {
        if (file.webkitRelativePath.includes('/Magnitude/') && file.name === `${selectedYear}_${selectedMonth}.xlsx`) {
            selectedFile = file;
            break;
        }
    }

    if (selectedFile) {
        const entities = await extractEntitiesFromExcel(selectedFile);
        populateEntityDropdown(entities);
    }
    
    hideLoader();
}

async function extractEntitiesFromExcel(file) {
    const workbook = await parseExcelFile(file);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const entities = new Set();
    jsonData.forEach(row => {
        if (row.length > 1 && row[1] && row[1].length === 6 && !isNaN(row[1][5])) {
            entities.add(row[1]);
        }
    });

    return Array.from(entities).sort();
}

function populateEntityDropdown(entities) {
    const entitySelect = document.getElementById('entity-select');
    entitySelect.innerHTML = '<option value="">--Choisir une entité--</option>';
    entities.forEach(entity => {
        const option = document.createElement('option');
        option.value = entity;
        option.textContent = entity;
        entitySelect.appendChild(option);
    });
}

async function handleEntitySelect() {
    const selectedEntity = document.getElementById('entity-select').value;
    const selectedYear = document.getElementById('year-select').value;
    const selectedMonth = document.getElementById('month-select').value;
    
    if (!selectedEntity || !selectedYear || !selectedMonth) return;
    
    showLoader('Chargement des données...');
    await new Promise(resolve => setTimeout(resolve, 50));
    
    const files = document.getElementById('folder-input').files;
    const previousYear = (parseInt(selectedYear) - 1).toString();
    const previousYear2 = (parseInt(selectedYear) - 2).toString();

    try {
        // Step 1: Magnitude
        setLoaderStep('magnitude');
        const [currentYearData, decemberPreviousYearData, previousYearData] = await Promise.all([
            getEntityDataForYear(files, selectedYear, selectedMonth, selectedEntity),
            getEntityDataForYear(files, previousYear, '12', selectedEntity),
            getEntityDataForYear(files, previousYear, selectedMonth, selectedEntity)
        ]);
        
        // Step 2: Kyriba
        setLoaderStep('kyriba');
        const [kyribaFile, kyribaFilePreviousYear, kyribaFilePreviousYear2] = await Promise.all([
            getKyribaFileForYear(selectedYear),
            getKyribaFileForYear(previousYear),
            getKyribaFileForYear(previousYear2)
        ]);
        
        // Step 3: Process
        setLoaderStep('process');
        const cashPromises = [
            kyribaFile ? extractDailyCashData(kyribaFile, selectedEntity) : Promise.resolve([]),
            kyribaFilePreviousYear ? extractDailyCashData(kyribaFilePreviousYear, selectedEntity) : Promise.resolve([]),
            kyribaFilePreviousYear2 ? extractDailyCashData(kyribaFilePreviousYear2, selectedEntity) : Promise.resolve([])
        ];
        const [dailyCashData, dailyCashDataPreviousYear, dailyCashDataPreviousYear2] = await Promise.all(cashPromises);
        
        // Step 4: Display
        setLoaderStep('display');
        await new Promise(resolve => setTimeout(resolve, 100));
        
        if (currentYearData) {
            displayDataForEntity(
                currentYearData.file, 
                currentYearData.entityData, 
                currentYearData.headers, 
                decemberPreviousYearData, 
                previousYearData
            );
        }
        
        displayCashChart(dailyCashData, dailyCashDataPreviousYear, dailyCashDataPreviousYear2, selectedYear);
        
    } catch (error) {
        console.error('Erreur lors du chargement:', error);
        alert('Une erreur est survenue lors du chargement des données.');
    }
    
    hideLoader();
}

// ═══════════════════════════════════════════════════════════════════════════
//                              PARTIE MAGNITUDE
// ═══════════════════════════════════════════════════════════════════════════

async function getEntityDataForYear(files, year, month, entity) {
    let selectedFile = null;

    for (const file of files) {
        if (file.webkitRelativePath.includes('/Magnitude/') && file.name === `${year}_${month}.xlsx`) {
            selectedFile = file;
            break;
        }
    }

    if (!selectedFile) return null;

    const workbook = await parseExcelFile(selectedFile);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = jsonData[0];
    let entityData = null;

    for (const row of jsonData.slice(1)) {
        if (row[1] === entity) {
            entityData = row;
            break;
        }
    }

    if (entityData) {
        return {
            file: selectedFile,
            headers: [...headers],
            entityData: [...entityData]
        };
    }
    return null;
}

async function displayDataForEntity(file, entityData, headers, decemberPreviousYearData, previousYearData) {
    const codeIndex = headers.indexOf('CODE');
    if (codeIndex !== -1) {
        headers.splice(codeIndex, 1);
        entityData.splice(codeIndex, 1);
    }
    const dateIndex = headers.indexOf('DATE');
    if (dateIndex !== -1) {
        headers.splice(dateIndex, 1);
        entityData.splice(dateIndex, 1);
    }

    headers[0] = entityData[0];
    entityData[0] = file.name.replace('.xlsx', '');

    let decemberEntityData = null;
    if (decemberPreviousYearData) {
        const decemberHeaders = [...decemberPreviousYearData.headers];
        decemberEntityData = [...decemberPreviousYearData.entityData];
        
        const decCodeIndex = decemberHeaders.indexOf('CODE');
        if (decCodeIndex !== -1) {
            decemberHeaders.splice(decCodeIndex, 1);
            decemberEntityData.splice(decCodeIndex, 1);
        }
        const decDateIndex = decemberHeaders.indexOf('DATE');
        if (decDateIndex !== -1) {
            decemberHeaders.splice(decDateIndex, 1);
            decemberEntityData.splice(decDateIndex, 1);
        }
        decemberEntityData[0] = decemberPreviousYearData.file.name.replace('.xlsx', '');
    }

    let previousEntityData = null;
    if (previousYearData) {
        const previousHeaders = [...previousYearData.headers];
        previousEntityData = [...previousYearData.entityData];
        
        const prevCodeIndex = previousHeaders.indexOf('CODE');
        if (prevCodeIndex !== -1) {
            previousHeaders.splice(prevCodeIndex, 1);
            previousEntityData.splice(prevCodeIndex, 1);
        }
        const prevDateIndex = previousHeaders.indexOf('DATE');
        if (prevDateIndex !== -1) {
            previousHeaders.splice(prevDateIndex, 1);
            previousEntityData.splice(prevDateIndex, 1);
        }
        previousEntityData[0] = previousYearData.file.name.replace('.xlsx', '');
    }

    displayTransposedData(headers, entityData, decemberEntityData, previousEntityData);
}

function formatNumberWithSpaces(number) {
    return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

function displayTransposedData(headers, data, decemberData, previousYearData) {
    const tableBody = document.querySelector('#data-table tbody');
    tableBody.innerHTML = '';

    const fragment = document.createDocumentFragment();

    headers.forEach((header, index) => {
        const row = document.createElement('tr');
        let keyCell = document.createElement('th');
        let valueCell = document.createElement('td');
        let decemberValueCell = document.createElement('td');
        let previousValueCell = document.createElement('td');

        if (index === 0) {
            keyCell = document.createElement('main-tr');
            valueCell = document.createElement('main-tr');
            decemberValueCell = document.createElement('main-tr');
            previousValueCell = document.createElement('main-tr');
        }

        if (parseFloat(data[index]) < 0) {
            valueCell.classList.add('negative');
        }

        if (headers[index] === 'CFS FROM EBIT') {
            keyCell = document.createElement('head-cfs');
            valueCell = document.createElement('head-cfs');
            decemberValueCell = document.createElement('head-cfs');
            previousValueCell = document.createElement('head-cfs');
        }

        if (headers[index] === 'NET CASH VARIATION') {
            keyCell = document.createElement('total-tr');
            valueCell = document.createElement('total-td');
            decemberValueCell = document.createElement('total-td');
            previousValueCell = document.createElement('total-td');
            valueCell.classList.remove('negative');
        }

        let value = data[index];
        if (value !== undefined && !isNaN(value) && value !== '') {
            value = formatNumberWithSpaces(parseFloat(value).toFixed(0));
        } else {
            value = data[index];
        }

        let decemberValue = '';
        if (decemberData && decemberData[index] !== undefined) {
            decemberValue = decemberData[index];
            if (!isNaN(decemberValue) && decemberValue !== '') {
                decemberValue = formatNumberWithSpaces(parseFloat(decemberValue).toFixed(0));
            }
            if (parseFloat(decemberData[index]) < 0) {
                decemberValueCell.classList.add('negative');
            }
        }

        let previousValue = '';
        if (previousYearData && previousYearData[index] !== undefined) {
            previousValue = previousYearData[index];
            if (!isNaN(previousValue) && previousValue !== '') {
                previousValue = formatNumberWithSpaces(parseFloat(previousValue).toFixed(0));
            }
            if (parseFloat(previousYearData[index]) < 0) {
                previousValueCell.classList.add('negative');
            }
        }

        if (header === 'O.G' || header === '% OVERDUES' || header === 'EBIT OP/GI') {
            value = `${value} %`;
            if (decemberValue !== '') decemberValue = `${decemberValue} %`;
            if (previousValue !== '') previousValue = `${previousValue} %`;
        }
        
        if (header === '' && index > 0) {
            keyCell = document.createElement('blank-tr');
            valueCell = document.createElement('blank-tr');
            decemberValueCell = document.createElement('blank-tr');
            previousValueCell = document.createElement('blank-tr');
        }

        keyCell.textContent = header;
        valueCell.textContent = value || '';
        decemberValueCell.textContent = decemberValue || '';
        previousValueCell.textContent = previousValue || '';

        row.appendChild(keyCell);
        row.appendChild(valueCell);
        row.appendChild(decemberValueCell);
        row.appendChild(previousValueCell);
        fragment.appendChild(row);
    });

    tableBody.appendChild(fragment);
}

// ═══════════════════════════════════════════════════════════════════════════
//                              PARTIE KYRIBA
// ═══════════════════════════════════════════════════════════════════════════

async function getKyribaFileForYear(year) {
    const files = document.getElementById('folder-input').files;
    for (const file of files) {
        if (file.webkitRelativePath.includes('/Kyriba/') && file.name === `${year}.xlsx`) {
            return file;
        }
    }
    return null;
}

async function extractDailyCashData(file, entity) {
    const workbook = await parseExcelFile(file);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = jsonData[0];
    const totalentity = {};

    jsonData.forEach(row => {
        if (row[1] === entity) {
            const multiplier = row[0]?.slice(0, 1) === 'Z' ? -1 : 1;
            row.slice(2).forEach((value, index) => {
                const date = headers[index + 2];
                const cash = parseFloat(value) * multiplier || 0;
                totalentity[date] = (totalentity[date] || 0) + cash;
            });
        }
    });

    return Object.keys(totalentity).map(date => ({
        date: date,
        cash: totalentity[date]
    }));
}

function displayCashChart(dailyCashData, dailyCashDataPreviousYear, dailyCashDataPreviousYear2, selectedYear) {
    const ctx = document.getElementById('cashChart').getContext('2d');

    if (cashChart instanceof Chart) {
        cashChart.destroy();
    }

    cashChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: dailyCashData.map(data => data.date),
            datasets: [
                {
                    label: `Daily Cash (${selectedYear})`,
                    data: dailyCashData.map(data => data.cash),
                    borderColor: '#e47a38',
                    borderWidth: 2,
                    fill: false,
                    pointRadius: 0
                },
                {
                    label: `Daily Cash (${selectedYear - 1})`,
                    data: dailyCashDataPreviousYear.map(data => data.cash),
                    borderColor: '#38a1e4',
                    borderWidth: 2,
                    fill: false,
                    pointRadius: 0
                },
                {
                    label: `Daily Cash (${selectedYear - 2})`,
                    data: dailyCashDataPreviousYear2.map(data => data.cash),
                    borderColor: '#C75AD1',
                    borderWidth: 2,
                    fill: false,
                    pointRadius: 0
                }
            ]
        },
        options: {
            animation: {
                duration: 300
            },
            interaction: {
                mode: 'index',
                intersect: false
            },
            scales: {
                x: { beginAtZero: true },
                y: { beginAtZero: false }
            }
        }
    });
}