// Global variables to store data and charts
let tariffData = null;
let mainChart = null;
let resultsChart = null;
let latestYear = 2024; // Default value, will be updated from data

// DOM elements
let elasticityInput;
let tariffInput;
let elasticityValidation;
let tariffValidation;

// Sample data in case the Excel file cannot be loaded
const sampleData = [
    { year: 2020, dutyrate: 2.10, importshare: 14.5, duties: 74.4, imports: 3542.9 },
    { year: 2021, dutyrate: 2.20, importshare: 15.1, duties: 85.6, imports: 3889.1 },
    { year: 2022, dutyrate: 2.25, importshare: 15.3, duties: 93.2, imports: 4142.2 },
    { year: 2023, dutyrate: 2.30, importshare: 15.5, duties: 98.7, imports: 4291.3 },
    { year: 2024, dutyrate: 2.34, importshare: 15.7, duties: 104.5, imports: 4465.8 }
];

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    // Initialize DOM elements
    elasticityInput = document.getElementById('elasticity');
    tariffInput = document.getElementById('tariff');
    elasticityValidation = document.getElementById('elasticity-validation');
    tariffValidation = document.getElementById('tariff-validation');

    // Try to load the Excel file automatically
    tryLoadExcelFile();
});

// Try to load the Excel file automatically
async function tryLoadExcelFile() {
    try {
        // Try to load the Excel file
        tariffData = await loadExcelData('utoandGDP.xlsx');
        latestYear = Math.max(...tariffData.map(row => row.year));
        initializeCharts();
    } catch (error) {
        console.warn('Could not automatically load Excel file:', error);
        console.log('Using sample data instead.');

        // Use sample data as fallback
        tariffData = [...sampleData];
        latestYear = Math.max(...tariffData.map(row => row.year));
        initializeCharts();
    }
}

// Initialize charts after data is loaded
function initializeCharts() {
    // Remove any existing charts
    if (mainChart) mainChart.destroy();
    if (resultsChart) resultsChart.destroy();

    // Create the charts
    createMainChart();
    createResultsChart();

    // Set up event listeners for inputs
    setupInputListeners();
}

// Load and process Excel data from URL
async function loadExcelData(filePath) {
    const response = await fetch(filePath);
    const arrayBuffer = await response.arrayBuffer();

    // Parse the Excel file
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    // Assume the first sheet contains our data
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Convert to JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    // Process and return the data
    return jsonData.map(row => ({
        year: row.year,
        dutyrate: row.dutyrate,
        importshare: row.importshare,
        duties: row.duties,
        imports: row.imports
    }));
}

// Create the main chart (year vs dutyrate and importshare)
function createMainChart() {
    const ctx = document.getElementById('mainChart').getContext('2d');

    // Prepare data for the chart
    const years = tariffData.map(row => row.year);
    const dutyRates = tariffData.map(row => row.dutyrate);
    const importShares = tariffData.map(row => row.importshare);

    // Create the chart
    mainChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: years,
            datasets: [
                {
                    label: 'Duty Rate (%)',
                    data: dutyRates,
                    borderColor: '#0056b3',
                    backgroundColor: 'rgba(0, 86, 179, 0.1)',
                    yAxisID: 'y',
                    tension: 0.1
                },
                {
                    label: 'Import Share (%)',
                    data: importShares,
                    borderColor: '#28a745',
                    backgroundColor: 'rgba(40, 167, 69, 0.1)',
                    yAxisID: 'y1',
                    tension: 0.1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false,
            },
            stacked: false,
            scales: {
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    title: {
                        display: true,
                        text: 'Duty Rate (%)'
                    }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    grid: {
                        drawOnChartArea: false,
                    },
                    title: {
                        display: true,
                        text: 'Import Share (%)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Year'
                    }
                }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.y !== null) {
                                label += context.parsed.y.toFixed(2);
                            }
                            return label;
                        }
                    }
                }
            }
        }
    });
}

// Create the results chart (bar chart with 4 bars)
function createResultsChart() {
    const ctx = document.getElementById('resultsChart').getContext('2d');

    // Get the latest year data
    const latestYearData = tariffData.find(row => row.year === latestYear);

    if (!latestYearData) {
        console.error(`No data found for year ${latestYear}`);
        return;
    }

    // Get input values
    const elasticity = parseFloat(elasticityInput.value);
    const tariffRate = parseFloat(tariffInput.value);

    // Calculate the values for the chart
    const currentImports = latestYearData.imports / 1e9; // Convert to billions
    const currentDuties = latestYearData.duties / 1e9; // Convert to billions

    // Calculate new values based on elasticity and tariff
    const currentTariffRate = latestYearData.dutyrate;
    const importsAfterPolicy = currentImports * Math.pow(currentTariffRate / tariffRate, elasticity);
    const dutiesAfterPolicy = importsAfterPolicy * (tariffRate / 100);

    // Create the chart
    resultsChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Imports 2024', 'Tariff revenue 2024', 'Imports after policy', 'Tariff revenue after policy'],
            datasets: [{
                data: [currentImports, currentDuties, importsAfterPolicy, dutiesAfterPolicy],
                backgroundColor: [
                    'rgba(0, 86, 179, 0.7)',
                    'rgba(40, 167, 69, 0.7)',
                    'rgba(0, 86, 179, 0.4)',
                    'rgba(40, 167, 69, 0.4)'
                ],
                borderColor: [
                    'rgba(0, 86, 179, 1)',
                    'rgba(40, 167, 69, 1)',
                    'rgba(0, 86, 179, 1)',
                    'rgba(40, 167, 69, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `$${context.parsed.y.toFixed(2)} billion`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Billions of Dollars'
                    }
                }
            }
        }
    });
}

// Update the results chart based on new input values
function updateResultsChart() {
    if (!resultsChart || !tariffData) return;

    // Validate inputs
    const elasticity = parseFloat(elasticityInput.value);
    const tariffRate = parseFloat(tariffInput.value);

    if (isNaN(elasticity) || elasticity < 0) {
        elasticityValidation.textContent = 'Please enter a non-negative number';
        return;
    } else {
        elasticityValidation.textContent = '';
    }

    if (isNaN(tariffRate) || tariffRate < 0) {
        tariffValidation.textContent = 'Please enter a non-negative number';
        return;
    } else {
        tariffValidation.textContent = '';
    }

    // Get the latest year data
    const latestYearData = tariffData.find(row => row.year === latestYear);

    if (!latestYearData) {
        console.error(`No data found for year ${latestYear}`);
        return;
    }

    // Calculate the values for the chart
    const currentImports = latestYearData.imports / 1e9; // Convert to billions
    const currentDuties = latestYearData.duties / 1e9; // Convert to billions

    // Calculate new values based on elasticity and tariff
    const currentTariffRate = latestYearData.dutyrate;
    const importsAfterPolicy = currentImports * Math.pow(currentTariffRate / tariffRate, elasticity);
    const dutiesAfterPolicy = importsAfterPolicy * (tariffRate / 100);

    // Update the chart data
    resultsChart.data.datasets[0].data = [
        currentImports,
        currentDuties,
        importsAfterPolicy,
        dutiesAfterPolicy
    ];

    // Update the chart
    resultsChart.update();
}

// Set up event listeners for input fields
function setupInputListeners() {
    elasticityInput.addEventListener('input', () => {
        if (elasticityInput.value < 0) {
            elasticityInput.value = 0;
        }
        updateResultsChart();
    });

    tariffInput.addEventListener('input', () => {
        if (tariffInput.value < 0) {
            tariffInput.value = 0;
        }
        updateResultsChart();
    });
}
