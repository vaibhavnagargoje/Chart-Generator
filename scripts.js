// Global variables
let excelData = null;
let sheets = [];
let currentSheet = null;
let columns = [];
let chartInstance = null;
let selectedChartType = null;

// Initialize when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', function() {
    // Setup event listeners
    setupEventListeners();
});

// Set up event listeners
function setupEventListeners() {
    // File upload handling
    const fileInput = document.getElementById('excelFile');
    const fileLabel = document.querySelector('.file-upload-label');
    
    fileInput.addEventListener('change', handleFileUpload);
    
    // Drag and drop handling for file upload
    fileLabel.addEventListener('dragover', function(e) {
        e.preventDefault();
        fileLabel.classList.add('dragover');
    });
    
    fileLabel.addEventListener('dragleave', function() {
        fileLabel.classList.remove('dragover');
    });
    
    fileLabel.addEventListener('drop', function(e) {
        e.preventDefault();
        fileLabel.classList.remove('dragover');
        
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileUpload();
        }
    });
    
    // Sheet selection change
    document.getElementById('sheetSelect').addEventListener('change', function() {
        const sheetName = this.value;
        processSheet(sheetName);
    });
    
    // Add another Y-axis series
    document.getElementById('addSeries').addEventListener('click', addYAxisSelector);
    
    // Apply data range button
    document.getElementById('applyRange').addEventListener('click', updateDataPreview);
    
    // Chart type selection
    document.querySelectorAll('.chart-type-card').forEach(card => {
        card.addEventListener('click', function() {
            document.querySelectorAll('.chart-type-card').forEach(c => c.classList.remove('selected'));
            this.classList.add('selected');
            selectedChartType = this.getAttribute('data-type');
        });
    });
    
    // Generate chart button
    document.getElementById('generateChartBtn').addEventListener('click', generateChart);
    
    // Chart customization buttons
    document.getElementById('applyChartTitle').addEventListener('click', updateChartTitle);
    document.getElementById('applyAxisLabels').addEventListener('click', updateAxisLabels);
    
    // Export buttons
    document.getElementById('downloadImageBtn').addEventListener('click', downloadChartAsImage);
    document.getElementById('downloadCodeBtn').addEventListener('click', downloadChartCode);
    document.getElementById('copyDataBtn').addEventListener('click', copyChartData);
    
    // Add event listeners for filter controls
    document.getElementById('filterColumn').addEventListener('change', populateFilterValues);
    document.getElementById('filterValue').addEventListener('change', updateDataPreview);
    
    // Add event listener for chart filter column
    document.getElementById('filterColumn2').addEventListener('change', function() {
        // Store selected value for later use when generating chart
        updateDataPreview();
    });
    
    // Add event listener for chart filter value dropdown
    document.getElementById('chartFilterValue').addEventListener('change', function() {
        if (chartInstance) {
            // Update chart with new filter value without resetting
            updateChartWithFilter();
        }
    });
}

// Handle Excel file upload
function handleFileUpload() {
    const fileInput = document.getElementById('excelFile');
    const fileName = document.getElementById('file-name');
    const loadingIndicator = document.getElementById('loading-indicator');
    
    if (fileInput.files.length === 0) {
        return;
    }
    
    const file = fileInput.files[0];
    fileName.textContent = file.name;
    
    // Show loading indicator
    loadingIndicator.classList.remove('hidden');
    
    // Read the file
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            // Parse Excel file
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Store workbook data
            excelData = workbook;
            sheets = workbook.SheetNames;
            
            // Populate sheet selector
            populateSheetSelector(sheets);
            
            // Process the first sheet
            if (sheets.length > 0) {
                processSheet(sheets[0]);
            }
            
            // Show data selection section
            document.getElementById('data-selection').classList.remove('hidden');
            document.getElementById('chart-type-selection').classList.remove('hidden');
            
        } catch (error) {
            console.error('Error processing Excel file:', error);
            alert('Error processing the Excel file. Please make sure it\'s a valid Excel file.');
        } finally {
            // Hide loading indicator
            loadingIndicator.classList.add('hidden');
        }
    };
    
    reader.onerror = function() {
        alert('Error reading the file');
        loadingIndicator.classList.add('hidden');
    };
    
    reader.readAsArrayBuffer(file);
}

// Populate sheet selector dropdown
function populateSheetSelector(sheets) {
    const sheetSelect = document.getElementById('sheetSelect');
    sheetSelect.innerHTML = '';
    
    sheets.forEach(sheet => {
        const option = document.createElement('option');
        option.value = sheet;
        option.textContent = sheet;
        sheetSelect.appendChild(option);
    });
}

// Process a sheet from the workbook
function processSheet(sheetName) {
    if (!excelData) return;
    
    const worksheet = excelData.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    currentSheet = {
        name: sheetName,
        data: jsonData
    };
    
    // Extract columns from the first row
    if (jsonData.length > 0) {
        columns = [];
        const headerRow = jsonData[0];
        
        headerRow.forEach((header, index) => {
            const colLetter = XLSX.utils.encode_col(index);
            columns.push({
                index: index,
                letter: colLetter,
                name: header || `Column ${colLetter}`
            });
        });
        
        // Set total rows in end row input
        document.getElementById('endRow').value = jsonData.length;
        
        // Populate column selectors
        populateColumnSelectors();
        
        // Update data preview
        updateDataPreview();
    }
}

// Populate X and Y axis column selectors
function populateColumnSelectors() {
    const xAxisSelect = document.getElementById('xAxisSelect');
    const yAxisSelects = document.querySelectorAll('.y-axis-select');
    const filterColumn = document.getElementById('filterColumn');
    const filterColumn2 = document.getElementById('filterColumn2');
    
    // Clear existing options
    xAxisSelect.innerHTML = '';
    yAxisSelects.forEach(select => select.innerHTML = '');
    filterColumn.innerHTML = '<option value="">No Filter</option>'; // Reset with default option
    filterColumn2.innerHTML = '<option value="">No Chart Filter</option>'; // Reset with default option
    
    // Add options for each column
    columns.forEach(column => {
        const option = document.createElement('option');
        option.value = column.index;
        option.textContent = `${column.letter}: ${column.name}`;
        
        xAxisSelect.appendChild(option.cloneNode(true));
        filterColumn.appendChild(option.cloneNode(true)); // Add to filter dropdown
        filterColumn2.appendChild(option.cloneNode(true)); // Add to second filter dropdown
        
        yAxisSelects.forEach(select => {
            select.appendChild(option.cloneNode(true));
        });
    });
    
    // Set default selections (first column for X, second for Y)
    if (columns.length > 0) {
        xAxisSelect.value = 0; // First column as X-axis
        
        if (columns.length > 1) {
            yAxisSelects[0].value = 1; // Second column as Y-axis
        }
    }
    
    // Set up event listeners for preview updates
    xAxisSelect.addEventListener('change', updateDataPreview);
    yAxisSelects.forEach(select => {
        select.addEventListener('change', updateDataPreview);
    });
    
    // Add event listener for filter column dropdown
    filterColumn.addEventListener('change', populateFilterValues);
    document.getElementById('filterValue').addEventListener('change', updateDataPreview);
}

// New function to populate filter values based on selected column
function populateFilterValues() {
    const filterColumn = document.getElementById('filterColumn');
    const filterValue = document.getElementById('filterValue');
    
    // Reset filter value dropdown
    filterValue.innerHTML = '';
    
    if (!filterColumn.value) {
        // No filter column selected
        filterValue.disabled = true;
        filterValue.innerHTML = '<option value="">Select column first</option>';
        updateDataPreview();
        return;
    }
    
    // Enable the filter value dropdown
    filterValue.disabled = false;
    
    // Get the selected column index
    const columnIndex = parseInt(filterColumn.value);
    
    // Get unique values from this column
    const uniqueValues = new Set();
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
        const row = currentSheet.data[i];
        if (row && row[columnIndex] !== undefined) {
            uniqueValues.add(row[columnIndex]);
        }
    }
    
    // Add "All" option
    const allOption = document.createElement('option');
    allOption.value = "";
    allOption.textContent = "All Values";
    filterValue.appendChild(allOption);
    
    // Sort the unique values (strings alphabetically, numbers numerically)
    const sortedValues = Array.from(uniqueValues).sort((a, b) => {
        if (typeof a === 'number' && typeof b === 'number') {
            return a - b;
        }
        return String(a).localeCompare(String(b));
    });
    
    // Add options for each unique value
    sortedValues.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        filterValue.appendChild(option);
    });
    
    // Update preview to reflect the filtering
    updateDataPreview();
}

// Add a new Y-axis selector
function addYAxisSelector() {
    const yAxisSelectors = document.getElementById('yAxisSelectors');
    const colorIndex = yAxisSelectors.children.length;
    
    const yAxisItem = document.createElement('div');
    yAxisItem.className = 'y-axis-item';
    
    const select = document.createElement('select');
    select.className = 'y-axis-select';
    
    // Add options for each column
    columns.forEach(column => {
        const option = document.createElement('option');
        option.value = column.index;
        option.textContent = `${column.letter}: ${column.name}`;
        select.appendChild(option);
    });
    
    // Set a default selection (try to pick a different column if available)
    if (columns.length > colorIndex + 1) {
        select.value = colorIndex + 1;
    }
    
    const colorInput = document.createElement('input');
    colorInput.type = 'color';
    colorInput.className = 'series-color';
    colorInput.value = getColorFromPalette(colorIndex);
    
    const removeButton = document.createElement('button');
    removeButton.className = 'remove-y-axis';
    removeButton.textContent = '✕';
    removeButton.title = 'Remove series';
    removeButton.addEventListener('click', function() {
        yAxisSelectors.removeChild(yAxisItem);
        updateDataPreview();
    });
    
    yAxisItem.appendChild(select);
    yAxisItem.appendChild(colorInput);
    yAxisItem.appendChild(removeButton);
    
    yAxisSelectors.appendChild(yAxisItem);
    
    // Add change event for preview update
    select.addEventListener('change', updateDataPreview);
    colorInput.addEventListener('change', updateDataPreview);
}

// Get a color from predefined palette
function getColorFromPalette(index) {
    const colorPalette = [
        '#4e73df', '#1cc88a', '#36b9cc', '#f6c23e', '#e74a3b',
        '#6f42c1', '#5a5c69', '#858796', '#4287f5', '#41e169'
    ];
    
    return colorPalette[index % colorPalette.length];
}

// Helper function to safely compare values of any type
function valueMatches(cellValue, filterValue) {
    // Handle nulls and undefined
    if (cellValue === null || cellValue === undefined) {
        return filterValue === "null" || filterValue === "undefined" || filterValue === "";
    }
    
    // Convert to strings and compare (case-insensitive)
    const cellStr = String(cellValue).trim().toLowerCase();
    const filterStr = String(filterValue).trim().toLowerCase();
    
    return cellStr === filterStr;
}

// Improved validation for data range inputs
function validateDataRange() {
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    if (!currentSheet || !currentSheet.data) {
        return false;
    }
    
    // Validate range boundaries
    if (isNaN(startRow) || isNaN(endRow) || 
        startRow < 2 || // Header row is 1
        startRow > currentSheet.data.length || 
        endRow < startRow || 
        endRow > currentSheet.data.length) {
        
        alert('Invalid data range. Please check your start and end row values.');
        return false;
    }
    
    return true;
}

// Update data preview based on selected columns and range
function updateDataPreview() {
    if (!currentSheet || !currentSheet.data || currentSheet.data.length === 0) {
        return;
    }
    
    if (!validateDataRange()) {
        return;
    }
    
    const xAxisSelect = document.getElementById('xAxisSelect');
    const yAxisSelects = document.querySelectorAll('.y-axis-select');
    const xAxisPreview = document.getElementById('xAxisPreview');
    const yAxisPreview = document.getElementById('yAxisPreview');
    
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    // Get filter settings
    const filterColumnSelect = document.getElementById('filterColumn');
    const filterValueSelect = document.getElementById('filterValue');
    
    const useFilter = filterColumnSelect.value !== '';
    const filterColumnIndex = useFilter ? parseInt(filterColumnSelect.value) : -1;
    const filterValue = filterValueSelect.value;
    
    // Function to check if a row matches the filter with improved type handling
    function rowMatchesFilter(row) {
        if (!useFilter || filterValue === '') {
            return true; // No filter applied
        }
        
        return valueMatches(row[filterColumnIndex], filterValue);
    }
    
    // Get selected columns
    const xColIndex = parseInt(xAxisSelect.value);
    
    // Clear previews
    xAxisPreview.innerHTML = '';
    yAxisPreview.innerHTML = '';
    
    // Generate X-axis preview with filtering
    const xPreviewData = [];
    let totalRows = 0;
    let filteredRows = 0;
    
    // Optimize for large datasets by chunking the processing
    const chunkSize = 1000; // Process 1000 rows at a time
    
    for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
        const row = currentSheet.data[i];
        
        // Skip processing every few rows if we're dealing with a very large dataset
        if (endRow - startRow > 10000 && i % 10 !== 0) continue;
        
        totalRows++;
        
        if (row && row[xColIndex] !== undefined) {
            // Apply filter
            if (!rowMatchesFilter(row)) continue;
            
            filteredRows++;
            xPreviewData.push(row[xColIndex]);
        }
    }
    
    // Generate Y-axis preview for each selected Y column with filtering
    const yPreviewData = {};
    
    yAxisSelects.forEach((select, index) => {
        const yColIndex = parseInt(select.value);
        const columnName = columns[yColIndex].name;
        
        yPreviewData[columnName] = [];
        
        for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
            const row = currentSheet.data[i];
            
            // Skip processing every few rows if we're dealing with a very large dataset
            if (endRow - startRow > 10000 && i % 10 !== 0) continue;
            
            if (row && row[yColIndex] !== undefined) {
                // Apply filter
                if (!rowMatchesFilter(row)) continue;
                
                yPreviewData[columnName].push(row[yColIndex]);
            }
        }
    });

    // Display filter information if active
    let filterInfoHTML = '';
    if (useFilter) {
        const columnName = columns[filterColumnIndex]?.name || 'Unknown';
        filterInfoHTML = `<div class="filter-info">
            <strong>Filter:</strong> ${columnName} = ${filterValue || 'All Values'}<br>
            <strong>Matching rows:</strong> ${filteredRows} of ${totalRows} rows
        </div>`;
    }
    
    // Display X-axis preview (show first 10 items + count)
    const xPreviewHTML = document.createElement('div');
    xPreviewHTML.innerHTML = `<strong>First 10 values (${xPreviewData.length} total):</strong><br>`;
    
    if (filterInfoHTML) {
        xPreviewHTML.innerHTML += filterInfoHTML;
    }
    
    const xPreviewList = document.createElement('ul');
    xPreviewData.slice(0, 10).forEach(value => {
        const li = document.createElement('li');
        li.textContent = value;
        xPreviewList.appendChild(li);
    });
    
    if (xPreviewData.length > 10) {
        const li = document.createElement('li');
        li.textContent = `... and ${xPreviewData.length - 10} more`;
        xPreviewList.appendChild(li);
    }
    
    xPreviewHTML.appendChild(xPreviewList);
    xAxisPreview.appendChild(xPreviewHTML);
    
    // Display Y-axis preview
    const yPreviewHTML = document.createElement('div');
    yPreviewHTML.innerHTML = '<strong>Selected Series:</strong><br>';
    
    const yPreviewList = document.createElement('ul');
    Object.keys(yPreviewData).forEach(columnName => {
        const values = yPreviewData[columnName];
        const li = document.createElement('li');
        
        // Calculate min, max, avg
        if (values.length > 0) {
            const min = Math.min(...values.filter(v => !isNaN(parseFloat(v))));
            const max = Math.max(...values.filter(v => !isNaN(parseFloat(v))));
            const sum = values.filter(v => !isNaN(parseFloat(v))).reduce((a, b) => a + parseFloat(b), 0);
            const avg = values.filter(v => !isNaN(parseFloat(v))).length > 0 ? 
                sum / values.filter(v => !isNaN(parseFloat(v))).length : 0;
            
            li.innerHTML = `<strong>${columnName}</strong> (${values.length} values)<br>
                           Min: ${min.toFixed(2)}, Max: ${max.toFixed(2)}, Avg: ${avg.toFixed(2)}`;
        } else {
            li.innerHTML = `<strong>${columnName}</strong> (0 values)<br>No matching data for current filter`;
        }
        
        yPreviewList.appendChild(li);
    });
    
    yPreviewHTML.appendChild(yPreviewList);
    yAxisPreview.appendChild(yPreviewHTML);
}

// Collect data for the chart from user selections
function collectChartData() {
    if (!validateDataRange()) {
        return null;
    }
    
    const xAxisSelect = document.getElementById('xAxisSelect');
    const yAxisSelects = document.querySelectorAll('.y-axis-select');
    const colorInputs = document.querySelectorAll('.series-color');
    
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    // Get filter settings
    const filterColumnSelect = document.getElementById('filterColumn');
    const filterValueSelect = document.getElementById('filterValue');
    
    const useFilter = filterColumnSelect.value !== '';
    const filterColumnIndex = useFilter ? parseInt(filterColumnSelect.value) : -1;
    const filterValue = filterValueSelect.value;
    
    // Get chart filter column
    const filterColumn2 = document.getElementById('filterColumn2');
    const chartFilterColumnIndex = filterColumn2.value ? parseInt(filterColumn2.value) : -1;
    
    // Validate
    if (yAxisSelects.length === 0) {
        alert('At least one Y-axis series is required.');
        return null;
    }
    
    const xColIndex = parseInt(xAxisSelect.value);
    
    // Function to check if a row matches the filter with improved type handling
    function rowMatchesFilter(row) {
        if (!useFilter || filterValue === '') {
            return true; // No filter applied
        }
        
        return valueMatches(row[filterColumnIndex], filterValue);
    }
    
    // For single-series charts like pie, only use first Y series
    if (selectedChartType === 'pie' || selectedChartType === 'doughnut' || selectedChartType === 'polarArea') {
        if (yAxisSelects.length > 1) {
            alert(`${selectedChartType.charAt(0).toUpperCase() + selectedChartType.slice(1)} charts can only display one data series. Using the first selected series.`);
        }
        
        const yColIndex = parseInt(yAxisSelects[0].value);
        const seriesColor = colorInputs[0].value;
        
        // Collect labels and data
        const labels = [];
        const data = [];
        const backgroundColor = [];
        
        for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
            const row = currentSheet.data[i];
            if (row && row[xColIndex] !== undefined && row[yColIndex] !== undefined) {
                // Apply filter
                if (!rowMatchesFilter(row)) continue;
                
                // Handle numeric validation - only include valid numeric values
                const yValue = parseFloat(row[yColIndex]);
                if (!isNaN(yValue) && yValue !== 0) {
                    labels.push(String(row[xColIndex])); // Ensure labels are strings
                    data.push(yValue);
                    
                    // Generate distinct colors for each segment
                    const hue = (i * 137.5) % 360; // Use golden ratio for better color distribution
                    backgroundColor.push(`hsl(${hue}, 70%, 60%)`);
                }
            }
        }
        
        // Check if we have valid data
        if (data.length === 0) {
            alert('No valid data for this chart type. Please check your data selection and filter settings.');
            return null;
        }
        
        return {
            labels: labels,
            datasets: [{
                label: columns[yColIndex].name, // Add a label for the dataset
                data: data,
                backgroundColor: backgroundColor,
                borderColor: 'white',
                borderWidth: 1
            }]
        };
    }
    
    // For coordinate-based charts (scatter, bubble)
    if (selectedChartType === 'scatter' || selectedChartType === 'bubble') {
        const datasets = [];
        
        yAxisSelects.forEach((select, index) => {
            const yColIndex = parseInt(select.value);
            const seriesColor = colorInputs[index].value;
            const columnName = columns[yColIndex].name;
            
            const data = [];
            
            for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
                const row = currentSheet.data[i];
                if (row && row[xColIndex] !== undefined && row[yColIndex] !== undefined) {
                    // Apply filter
                    if (!rowMatchesFilter(row)) continue;
                    
                    const x = parseFloat(row[xColIndex]) || 0;
                    const y = parseFloat(row[yColIndex]) || 0;
                    
                    if (selectedChartType === 'bubble') {
                        // Use a third column for bubble size if available, otherwise use constant size
                        let size = 10;
                        if (row[yColIndex + 1] !== undefined) {
                            size = parseFloat(row[yColIndex + 1]) || 10;
                        }
                        data.push({ x, y, r: size });
                    } else {
                        data.push({ x, y });
                    }
                }
            }
            
            datasets.push({
                label: columnName,
                data: data,
                backgroundColor: seriesColor,
                borderColor: seriesColor,
                borderWidth: 1,
                pointRadius: 5,
                pointBackgroundColor: seriesColor
            });
        });
        
        return { datasets };
    }
    
    // For multi-series charts (bar, line, radar)
    const filteredRows = [];
    const xLabels = new Set();
    
    // First collect all rows that match the filter and their X values
    for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
        const row = currentSheet.data[i];
        if (row && row[xColIndex] !== undefined) {
            // Apply filter
            if (!rowMatchesFilter(row)) continue;
            
            filteredRows.push({ index: i, row: row });
            xLabels.add(row[xColIndex]);
        }
    }
    
    // Sort the X labels (numerically if possible, otherwise alphabetically)
    const labels = Array.from(xLabels).sort((a, b) => {
        const numA = parseFloat(a);
        const numB = parseFloat(b);
        
        if (!isNaN(numA) && !isNaN(numB)) {
            return numA - numB;
        }
        
        return String(a).localeCompare(String(b));
    });
    
    // Create datasets
    const datasets = [];
    
    // Collect data for each Y-axis series
    yAxisSelects.forEach((select, index) => {
        const yColIndex = parseInt(select.value);
        const seriesColor = colorInputs[index].value;
        const columnName = columns[yColIndex].name;
        
        const data = [];
        
        // For each label (x value), find the corresponding data
        labels.forEach(label => {
            // Find all rows with this X value
            const matchingRows = filteredRows.filter(item => item.row[xColIndex] === label);
            
            if (matchingRows.length > 0) {
                // Sum up all Y values for this X label
                let sum = 0;
                matchingRows.forEach(item => {
                    if (item.row[yColIndex] !== undefined) {
                        sum += parseFloat(item.row[yColIndex]) || 0;
                    }
                });
                data.push(sum);
            } else {
                data.push(0); // No data for this label
            }
        });
        
        const dataset = {
            label: columnName,
            data: data,
            backgroundColor: seriesColor,
            borderColor: seriesColor,
            borderWidth: 1
        };
        
        // Additional properties for specific chart types
        if (selectedChartType === 'line') {
            dataset.fill = false;
            dataset.tension = 0.1;
        } else if (selectedChartType === 'radar') {
            dataset.fill = true;
            dataset.backgroundColor = seriesColor + '50'; // Add transparency
        }
        
        datasets.push(dataset);
    });
    
    // If no data, alert the user
    if (labels.length === 0) {
        alert('No data matches your filter criteria. Please adjust your filter settings.');
        return null;
    }
    
    // When collecting chart data, make sure to store the original data rows for filtering
    if (chartFilterColumnIndex >= 0) {
        // Store row references for later filtering
        chartInstance.rowData = filteredRows.map(item => item.row);
    }
    
    return {
        labels: labels,
        datasets: datasets
    };
}

// Generate chart based on user selections
function generateChart() {
    if (!currentSheet || !selectedChartType) {
        alert('Please select a chart type and ensure you have data loaded.');
        return;
    }
    
    try {
        // Validate range
        if (!validateDataRange()) {
            return;
        }
        
        // Show loading message
        const loadingMessage = document.createElement('div');
        loadingMessage.className = 'chart-loading-message';
        loadingMessage.textContent = 'Processing data, please wait...';
        document.body.appendChild(loadingMessage);
        
        // Use setTimeout to allow the UI to update before heavy processing
        setTimeout(() => {
            try {
                const chartData = collectChartData();
                if (!chartData) {
                    document.body.removeChild(loadingMessage);
                    return;
                }
                
                // Destroy previous chart if exists
                if (chartInstance) {
                    chartInstance.destroy();
                }
                
                const ctx = document.getElementById('chartCanvas').getContext('2d');
                
                // Create chart configuration
                let chartType = selectedChartType;
                let processedData = chartData;
                
                // Handle special cases
                if (selectedChartType === 'stackedBar' || selectedChartType === 'percentStackedBar') {
                    chartType = 'bar'; // Chart.js uses 'bar' type with stacked option
                    
                    // For percentage stacked bars, convert data to percentages
                    if (selectedChartType === 'percentStackedBar') {
                        processedData = {
                            labels: chartData.labels,
                            datasets: calculatePercentageData(chartData.datasets, chartData.labels)
                        };
                    }
                }
                
                const chartConfig = {
                    type: chartType,
                    data: processedData,
                    options: generateChartOptions(selectedChartType)
                };
                
                // Add legend click handler for percentage stacked bar charts
                if (selectedChartType === 'percentStackedBar') {
                    // Store original data for recalculation
                    chartConfig.originalData = JSON.parse(JSON.stringify(chartData));
                    
                    // Add legend click handler
                    const legendClickHandler = {
                        id: 'legendClickHandler',
                        beforeInit: function(chart) {
                            // Save original handler reference
                            const originalLegendOnClick = chart.options.plugins.legend.onClick;
                            
                            // Override default legend click handler
                            chart.options.plugins.legend.onClick = function(e, legendItem, legend) {
                                // Call the original handler (toggles visibility)
                                originalLegendOnClick.call(this, e, legendItem, legend);
                                
                                // If it's a percent stacked bar, recalculate percentages
                                if (selectedChartType === 'percentStackedBar') {
                                    // Get currently visible datasets
                                    const visibleDatasets = chart.originalData.datasets.filter((dataset, i) => 
                                        !chart.getDatasetMeta(i).hidden
                                    );
                                    
                                    // Recalculate percentages based only on visible datasets
                                    const recalculatedDatasets = calculatePercentageData(
                                        visibleDatasets, 
                                        chart.originalData.labels
                                    );
                                    
                                    // Update each visible dataset with new percentages
                                    let visibleIndex = 0;
                                    chart.data.datasets.forEach((dataset, i) => {
                                        if (!chart.getDatasetMeta(i).hidden) {
                                            dataset.data = recalculatedDatasets[visibleIndex].data;
                                            visibleIndex++;
                                        }
                                    });
                                    
                                    chart.update();
                                }
                            };
                        }
                    };
                    
                    // Add the plugin to chart configuration
                    if (!chartConfig.plugins) chartConfig.plugins = [];
                    chartConfig.plugins.push(legendClickHandler);
                }
                
                // Create chart
                chartInstance = new Chart(ctx, chartConfig);
                
                // Store original data for percentage recalculation when toggling legend
                if (selectedChartType === 'percentStackedBar') {
                    chartInstance.originalData = chartConfig.originalData;
                }
                
                // Store original complete chart data for filtering
                chartInstance.fullData = JSON.parse(JSON.stringify(chartData));
                
                // Populate chart filter dropdown based on the selected filter column
                populateChartFilterValues();
                
                // Make sure the chart filter controls are visible if a chart filter column was selected
                const filterColumn2 = document.getElementById('filterColumn2');
                const chartFilterControls = document.querySelector('.chart-filter-controls');
                
                if (filterColumn2.value) {
                    chartFilterControls.style.display = 'flex';
                } else {
                    chartFilterControls.style.display = 'none';
                }
                
                // Remove loading message
                document.body.removeChild(loadingMessage);
                
                // Show chart display area
                document.getElementById('chart-display').classList.remove('hidden');
                
                // Set default titles
                document.getElementById('chartTitle').value = 'Excel Data Chart';
                document.getElementById('xAxisLabel').value = columns[parseInt(document.getElementById('xAxisSelect').value)].name;
                document.getElementById('yAxisLabel').value = selectedChartType === 'percentStackedBar' ? 'Percentage (%)' : 'Values';
            } catch (error) {
                // Remove loading message
                document.body.removeChild(loadingMessage);
                console.error('Error generating chart:', error);
                alert('Error generating chart: ' + error.message);
            }
        }, 50); // Small delay to let the UI update
        
    } catch (error) {
        console.error('Error preparing chart generation:', error);
        alert('Error preparing chart: ' + error.message);
    }
}

// Add a new function to calculate percentage data
function calculatePercentageData(datasets, labels) {
    // For each label/category, calculate the sum of all dataset values
    const totals = labels.map((_, labelIndex) => {
        return datasets.reduce((sum, dataset) => {
            const value = dataset.data[labelIndex] || 0;
            return sum + Math.abs(value);  // Use absolute value to handle negative numbers
        }, 0);
    });
    
    // Convert each dataset value to a percentage of the total
    return datasets.map(dataset => {
        const percentData = dataset.data.map((value, index) => {
            if (totals[index] === 0) return 0;  // Avoid division by zero
            return (Math.abs(value) / totals[index]) * 100;  // Calculate percentage
        });
        
        return {
            ...dataset,
            data: percentData
        };
    });
}

// Generate chart options based on chart type
function generateChartOptions(chartType) {
    const baseOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'top',
            },
            tooltip: {
                enabled: true
            },
            title: {
                display: true,
                text: 'Excel Data Chart',
                font: {
                    size: 18
                }
            }
        }
    };
    
    // Add chart-specific options
    switch (chartType) {
        case 'bar':
            return {
                ...baseOptions,
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: 'Categories'
                        }
                    },
                    y: {
                        display: true,
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Values'
                        }
                    }
                }
            };
            
        case 'stackedBar':
            return {
                ...baseOptions,
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: 'Categories'
                        },
                        stacked: true
                    },
                    y: {
                        display: true,
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Values'
                        },
                        stacked: true
                    }
                }
            };
            
        case 'line':
            return {
                ...baseOptions,
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: 'Categories'
                        }
                    },
                    y: {
                        display: true,
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Values'
                        }
                    }
                }
            };
            
        case 'pie':
        case 'doughnut':
            return {
                ...baseOptions,
                plugins: {
                    ...baseOptions.plugins,
                    legend: {
                        position: 'right',
                        labels: {
                            generateLabels: function(chart) {
                                // Generate better labels if we have many segments
                                const data = chart.data;
                                if (data.labels.length && data.datasets.length) {
                                    return data.labels.map(function(label, i) {
                                        const meta = chart.getDatasetMeta(0);
                                        const style = meta.controller.getStyle(i);
                                        
                                        return {
                                            text: label,
                                            fillStyle: style.backgroundColor,
                                            strokeStyle: style.borderColor,
                                            lineWidth: style.borderWidth,
                                            hidden: isNaN(data.datasets[0].data[i]) || meta.data[i].hidden,
                                            index: i
                                        };
                                    });
                                }
                                return [];
                            }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const value = context.formattedValue;
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = Math.round((context.raw / total) * 100);
                                return `${label}: ${value} (${percentage}%)`;
                            }
                        }
                    }
                }
            };
            
        case 'polarArea':
            return {
                ...baseOptions,
                plugins: {
                    ...baseOptions.plugins,
                    legend: {
                        position: 'right',
                    }
                }
            };
            
        case 'radar':
            return {
                ...baseOptions,
                scales: {
                    r: {
                        beginAtZero: true
                    }
                }
            };
            
        case 'scatter':
        case 'bubble':
            return {
                ...baseOptions,
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: 'X Values'
                        }
                    },
                    y: {
                        display: true,
                        title: {
                            display: true,
                            text: 'Y Values'
                        }
                    }
                }
            };
            
        case 'percentStackedBar':
            return {
                ...baseOptions,
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: 'Categories'
                        },
                        stacked: true
                    },
                    y: {
                        display: true,
                        stacked: true,
                        beginAtZero: true,
                        min: 0,
                        max: 100,
                        ticks: {
                            callback: function(value) {
                                return value + '%';
                            }
                        },
                        title: {
                            display: true,
                            text: 'Percentage'
                        }
                    }
                },
                plugins: {
                    ...baseOptions.plugins,
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const labelText = context.dataset.label || '';
                                const value = context.raw.toFixed(1);
                                return `${labelText}: ${value}%`;
                            }
                        }
                    }
                }
            };
            
        default:
            return baseOptions;
    }
}

// Update chart title
function updateChartTitle() {
    if (!chartInstance) return;
    
    const titleText = document.getElementById('chartTitle').value;
    
    chartInstance.options.plugins.title.text = titleText;
    chartInstance.update();
}

// Update axis labels
function updateAxisLabels() {
    if (!chartInstance) return;
    
    const xAxisText = document.getElementById('xAxisLabel').value;
    const yAxisText = document.getElementById('yAxisLabel').value;
    
    if (chartInstance.options.scales) {
        // For charts with standard x, y axes
        if (chartInstance.options.scales.x) {
            chartInstance.options.scales.x.title.display = true;
            chartInstance.options.scales.x.title.text = xAxisText;
        }
        
        if (chartInstance.options.scales.y) {
            chartInstance.options.scales.y.title.display = true;
            chartInstance.options.scales.y.title.text = yAxisText;
        }
        
        chartInstance.update();
    }
}

// Download chart as image
function downloadChartAsImage() {
    if (!chartInstance) {
        alert('Please generate a chart first.');
        return;
    }
    
    const canvas = document.getElementById('chartCanvas');
    const image = canvas.toDataURL('image/png');
    const downloadLink = document.createElement('a');
    downloadLink.href = image;
    downloadLink.download = 'chart.png';
    downloadLink.click();
}

// Copy chart data to clipboard
function copyChartData() {
    if (!chartInstance) {
        alert('Please generate a chart first.');
        return;
    }
    
    const data = chartInstance.data;
    const jsonString = JSON.stringify(data, null, 2);
    
    navigator.clipboard.writeText(jsonString).then(() => {
        alert('Chart data copied to clipboard!');
    }).catch(err => {
        console.error('Failed to copy chart data: ', err);
        alert('Failed to copy chart data. You may need to use a secure context (HTTPS).');
    });
}

// Download embeddable chart code
function downloadChartCode() {
    if (!chartInstance) {
        alert('Please generate a chart first.');
        return;
    }
    
    // Get chart configuration
    let chartType = selectedChartType;
    let chartData = chartInstance.data;
    let chartOptions = chartInstance.options;
    let additionalCode = '';
    
    // Handle special case for stacked bar and percentage stacked bar
    if (chartType === 'stackedBar' || chartType === 'percentStackedBar') {
        chartType = 'bar';
        
        // For percentage stacked bars, add the percentage calculation function
        if (selectedChartType === 'percentStackedBar') {
            additionalCode = `
        // Calculate percentage data for stacked bar
        function calculatePercentageData(datasets, labels) {
            // For each label/category, calculate the sum of all dataset values
            const totals = labels.map((_, labelIndex) => {
                return datasets.reduce((sum, dataset) => {
                    const value = dataset.data[labelIndex] || 0;
                    return sum + Math.abs(value);
                }, 0);
            });
            
            // Convert each dataset value to a percentage of the total
            return datasets.map(dataset => {
                const percentData = dataset.data.map((value, index) => {
                    if (totals[index] === 0) return 0;
                    return (Math.abs(value) / totals[index]) * 100;
                });
                
                return {
                    ...dataset,
                    data: percentData
                };
            });
        }
        
        // Store original data for recalculation when toggling legend items
        const originalData = ${JSON.stringify(chartInstance.originalData || chartData, null, 2)};
        
        // Override legend click handler to recalculate percentages
        const originalLegendOnClick = Chart.defaults.plugins.legend.onClick;
        Chart.defaults.plugins.legend.onClick = function(e, legendItem, legend) {
            // Call the original handler (toggles visibility)
            originalLegendOnClick.call(this, e, legendItem, legend);
            
            // Get currently visible datasets
            const visibleDatasets = originalData.datasets.filter((dataset, i) => 
                !chart.getDatasetMeta(i).hidden
            );
            
            // Recalculate percentages based only on visible datasets
            const recalculatedDatasets = calculatePercentageData(
                visibleDatasets, 
                originalData.labels
            );
            
            // Update each visible dataset with new percentages
            let visibleIndex = 0;
            chart.data.datasets.forEach((dataset, i) => {
                if (!chart.getDatasetMeta(i).hidden) {
                    dataset.data = recalculatedDatasets[visibleIndex].data;
                    visibleIndex++;
                }
            });
            
            chart.update();
        };
        
        // Convert data to percentages initially
        data.datasets = calculatePercentageData(data.datasets, data.labels);`;
        }
    }
    
    // Create HTML template for embeddable chart
    const htmlTemplate = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Embedded Chart</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
        }
        .chart-container {
            width: 100%;
            height: 100%;
            min-height: 400px;
            padding: 20px;
            box-sizing: border-box;
        }
    </style>
</head>
<body>
    <div class="chart-container">
        <canvas id="embeddedChart"></canvas>
    </div>
    
    <script>
        // Initialize chart when the page loads
        document.addEventListener('DOMContentLoaded', function() {
            const ctx = document.getElementById('embeddedChart').getContext('2d');
            
            // Chart data
            const data = ${JSON.stringify(chartData, null, 2)};
            
            // Chart options
            const options = ${JSON.stringify(chartOptions, null, 2)};
            ${additionalCode}
            
            // Create chart
            const chart = new Chart(ctx, {
                type: '${chartType}',
                data: data,
                options: options
            });
        });
    </script>
</body>
</html>`;
    
    // Create a Blob with the HTML content
    const blob = new Blob([htmlTemplate], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    
    // Create download link
    const downloadLink = document.createElement('a');
    downloadLink.href = url;
    downloadLink.download = 'embeddable-chart.html';
    
    // Trigger download
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    
    // Clean up
    URL.revokeObjectURL(url);
}

// Populate chart filter dropdown based on selected column
function populateChartFilterValues() {
    const filterColumn2 = document.getElementById('filterColumn2');
    const chartFilterValue = document.getElementById('chartFilterValue');
    const chartFilterLabel = document.getElementById('chartFilterLabel');
    
    // Clear existing options
    chartFilterValue.innerHTML = '';
    
    // Add "All Values" option
    const allOption = document.createElement('option');
    allOption.value = "";
    allOption.textContent = "All Values";
    chartFilterValue.appendChild(allOption);
    
    if (!filterColumn2.value) {
        // No filter column selected
        chartFilterValue.disabled = true;
        document.querySelector('.chart-filter-controls').style.display = 'none';
        return;
    }
    
    // Show the chart filter controls
    document.querySelector('.chart-filter-controls').style.display = 'flex';
    
    // Enable the filter value dropdown
    chartFilterValue.disabled = false;
    
    // Get the selected column index
    const columnIndex = parseInt(filterColumn2.value);
    const columnName = columns[columnIndex].name;
    
    // Update the label
    chartFilterLabel.textContent = `Filter by ${columnName}:`;
    
    // Use a Map to preserve the original data types
    const uniqueValuesMap = new Map();
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    // Apply main filter (if any)
    const filterColumnSelect = document.getElementById('filterColumn');
    const filterValueSelect = document.getElementById('filterValue');
    const useFilter = filterColumnSelect.value !== '';
    const filterColumnIndex = useFilter ? parseInt(filterColumnSelect.value) : -1;
    const filterValue = filterValueSelect.value;
    
    // Function to check if a row matches the main filter
    function rowMatchesMainFilter(row) {
        if (!useFilter || filterValue === '') {
            return true; // No filter applied
        }
        return valueMatches(row[filterColumnIndex], filterValue);
    }
    
    // Get values from current chart data first - this ensures we only show values that are actually in the chart
    if (chartInstance && chartInstance.fullData) {
        // For bar/line/radar charts
        if (chartInstance.fullData.labels && chartInstance.fullData.datasets) {
            // Find all rows that match our current chart data
            for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
                const row = currentSheet.data[i];
                if (!row) continue;
                
                const xColIndex = parseInt(document.getElementById('xAxisSelect').value);
                const rowXValue = row[xColIndex];
                
                // Check if this row's X value is in our chart labels
                if (chartInstance.fullData.labels.includes(rowXValue)) {
                    // Apply main filter
                    if (!rowMatchesMainFilter(row)) continue;
                    
                    if (row[columnIndex] !== undefined) {
                        // Use the value itself as key to preserve type
                        const value = row[columnIndex];
                        uniqueValuesMap.set(String(value), value);
                    }
                }
            }
        }
        // For scatter/bubble charts
        else if (chartInstance.fullData.datasets) {
            // Each point has x,y coordinates
            const points = [];
            chartInstance.fullData.datasets.forEach(dataset => {
                dataset.data.forEach(point => {
                    points.push(point);
                });
            });
            
            // Find matching rows
            for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i++) {
                const row = currentSheet.data[i];
                if (!row) continue;
                
                const xColIndex = parseInt(document.getElementById('xAxisSelect').value);
                const rowX = parseFloat(row[xColIndex]) || 0;
                
                // Check all Y axis columns
                const yAxisSelects = document.querySelectorAll('.y-axis-select');
                const yColIndices = Array.from(yAxisSelects).map(select => parseInt(select.value));
                
                let matchFound = false;
                for (const yColIndex of yColIndices) {
                    const rowY = parseFloat(row[yColIndex]) || 0;
                    
                    // Check if this point exists in our chart data
                    for (const point of points) {
                        const epsilon = 0.0001;
                        if (Math.abs(point.x - rowX) < epsilon && Math.abs(point.y - rowY) < epsilon) {
                            matchFound = true;
                            break;
                        }
                    }
                    
                    if (matchFound) break;
                }
                
                if (matchFound) {
                    // Apply main filter
                    if (!rowMatchesMainFilter(row)) continue;
                    
                    if (row[columnIndex] !== undefined) {
                        // Use the value itself as key to preserve type
                        const value = row[columnIndex];
                        uniqueValuesMap.set(String(value), value);
                    }
                }
            }
        }
    }
    
    // Sort the unique values (strings alphabetically, numbers numerically)
    const uniqueValues = Array.from(uniqueValuesMap.values());
    
    const sortedValues = uniqueValues.sort((a, b) => {
        if (typeof a === 'number' && typeof b === 'number') {
            return a - b;
        }
        return String(a).localeCompare(String(b));
    });
    
    // Add options for each unique value
    sortedValues.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        chartFilterValue.appendChild(option);
    });
    
    console.log(`Chart filter populated with ${sortedValues.length} unique values`);
}

// Update chart with new filter value
function updateChartWithFilter() {
    if (!chartInstance || !chartInstance.fullData) return;
    
    const chartFilterValue = document.getElementById('chartFilterValue');
    const filterColumn2 = document.getElementById('filterColumn2');
    
    if (!filterColumn2.value) return;
    
    const columnIndex = parseInt(filterColumn2.value);
    const filterValue = chartFilterValue.value;
    
    console.log(`Applying chart filter: Column index ${columnIndex}, Value: "${filterValue}"`);
    
    // Deep clone the full data
    const fullData = JSON.parse(JSON.stringify(chartInstance.fullData));
    
    if (filterValue === '') {
        // No filter, use full data
        console.log('Using full dataset (no filter)');
        updateChartWithData(fullData);
        return;
    }
    
    try {
        // For pie/doughnut/polarArea charts (single series)
        if (selectedChartType === 'pie' || selectedChartType === 'doughnut' || selectedChartType === 'polarArea') {
            const filteredLabels = [];
            const filteredData = [];
            const backgroundColor = [];
            
            // Match original data indexes with filter value
            for (let i = 0; i < fullData.labels.length; i++) {
                const label = fullData.labels[i];
                
                // Get all rows that match this label
                const matchingRows = findAllRowsByLabel(label, columnIndex);
                let matchesFilter = matchingRows.some(row => 
                    valueMatches(row[columnIndex], filterValue)
                );
                
                if (matchesFilter) {
                    filteredLabels.push(label);
                    filteredData.push(fullData.datasets[0].data[i]);
                    backgroundColor.push(fullData.datasets[0].backgroundColor[i]);
                }
            }
            
            // Make sure we have data to display
            if (filteredLabels.length === 0) {
                alert('No data matches your filter criteria.');
                return;
            }
            
            // Update chart with filtered data
            chartInstance.data.labels = filteredLabels;
            chartInstance.data.datasets[0].data = filteredData;
            chartInstance.data.datasets[0].backgroundColor = backgroundColor;
        }
        // For scatter/bubble charts
        else if (selectedChartType === 'scatter' || selectedChartType === 'bubble') {
            let hasData = false;
            
            chartInstance.data.datasets.forEach((dataset, datasetIndex) => {
                const filteredData = [];
                
                // Filter points based on the selected value
                for (let pointIndex = 0; pointIndex < fullData.datasets[datasetIndex].data.length; pointIndex++) {
                    const point = fullData.datasets[datasetIndex].data[pointIndex];
                    
                    // Find all rows that match these coordinates
                    const matchingRows = findAllRowsByCoordinates(point.x, point.y, columnIndex);
                    let matchesFilter = matchingRows.some(row => 
                        valueMatches(row[columnIndex], filterValue)
                    );
                    
                    if (matchesFilter) {
                        filteredData.push(point);
                        hasData = true;
                    }
                }
                
                // Update this dataset
                dataset.data = filteredData;
            });
            
            if (!hasData) {
                alert('No data matches your filter criteria.');
                return;
            }
        }
        // For other chart types (bar, line, radar, etc.)
        else {
            const filteredLabels = [];
            const filteredDatasets = fullData.datasets.map(dataset => ({
                ...dataset,
                data: []
            }));
            
            // For each label, check if any matching row has the filter value
            for (let labelIndex = 0; labelIndex < fullData.labels.length; labelIndex++) {
                const label = fullData.labels[labelIndex];
                
                // Find all rows that have this label
                const matchingRows = findAllRowsByLabel(label, columnIndex);
                
                // Check if any rows match the filter value
                const matchesFilter = matchingRows.some(row => 
                    valueMatches(row[columnIndex], filterValue)
                );
                
                if (matchesFilter) {
                    filteredLabels.push(label);
                    
                    // For each dataset, add the corresponding data point
                    fullData.datasets.forEach((dataset, datasetIndex) => {
                        filteredDatasets[datasetIndex].data.push(dataset.data[labelIndex]);
                    });
                }
            }
            
            // Make sure we have data to display
            if (filteredLabels.length === 0) {
                alert('No data matches your filter criteria.');
                return;
            }
            
            // Update chart data
            chartInstance.data.labels = filteredLabels;
            chartInstance.data.datasets.forEach((dataset, i) => {
                dataset.data = filteredDatasets[i].data;
            });
        }
        
        // Special cases for different chart types
        if (selectedChartType === 'percentStackedBar' && chartInstance.data.labels.length > 0) {
            // For percentage stacked bars, recalculate percentages
            const recalculatedDatasets = calculatePercentageData(
                chartInstance.data.datasets, 
                chartInstance.data.labels
            );
            
            chartInstance.data.datasets.forEach((dataset, i) => {
                dataset.data = recalculatedDatasets[i].data;
            });
        }
        
        // Update the chart
        chartInstance.update();
        console.log('Chart updated with filtered data');
        
    } catch (error) {
        console.error('Error filtering chart:', error);
        alert('Error applying filter. Please try again with different criteria.');
    }
}

// Helper function to find all rows with a specific label
function findAllRowsByLabel(labelValue, filterColumnIndex) {
    const xAxisSelect = document.getElementById('xAxisSelect');
    const xColIndex = parseInt(xAxisSelect.value);
    
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    // Get filter settings (for the first filter)
    const filterColumnSelect = document.getElementById('filterColumn');
    const filterValueSelect = document.getElementById('filterValue');
    const useFilter = filterColumnSelect.value !== '';
    const mainFilterColumnIndex = useFilter ? parseInt(filterColumnSelect.value) : -1;
    const mainFilterValue = filterValueSelect.value;
    
    const matchingRows = [];
    
    // Optimize for large datasets
    const checkEveryNthRow = endRow - startRow > 10000 ? 2 : 1;
    
    for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i += checkEveryNthRow) {
        const row = currentSheet.data[i];
        if (!row) continue;
        
        // Check if the X value matches our label
        if (valueMatches(row[xColIndex], labelValue)) {
            // Apply main filter if present
            if (useFilter && mainFilterValue !== '') {
                if (!valueMatches(row[mainFilterColumnIndex], mainFilterValue)) {
                    continue;
                }
            }
            
            matchingRows.push(row);
        }
    }
    
    return matchingRows;
}

// Helper function to find rows by x,y coordinates (for scatter/bubble charts)
function findAllRowsByCoordinates(x, y, filterColumnIndex) {
    const xAxisSelect = document.getElementById('xAxisSelect');
    const yAxisSelects = document.querySelectorAll('.y-axis-select');
    
    const xColIndex = parseInt(xAxisSelect.value);
    const yColIndices = Array.from(yAxisSelects).map(select => parseInt(select.value));
    
    const startRow = parseInt(document.getElementById('startRow').value);
    const endRow = parseInt(document.getElementById('endRow').value);
    
    // Get filter settings (for the first filter)
    const filterColumnSelect = document.getElementById('filterColumn');
    const filterValueSelect = document.getElementById('filterValue');
    const useFilter = filterColumnSelect.value !== '';
    const mainFilterColumnIndex = useFilter ? parseInt(filterColumnSelect.value) : -1;
    const mainFilterValue = filterValueSelect.value;
    
    const matchingRows = [];
    
    // Optimize for large datasets
    const checkEveryNthRow = endRow - startRow > 10000 ? 5 : 1;
    
    for (let i = startRow - 1; i < endRow && i < currentSheet.data.length; i += checkEveryNthRow) {
        const row = currentSheet.data[i];
        if (!row) continue;
        
        const rowX = parseFloat(row[xColIndex]) || 0;
        
        for (const yColIndex of yColIndices) {
            const rowY = parseFloat(row[yColIndex]) || 0;
            
            // Compare with a small epsilon for floating point comparison
            const epsilon = 0.0001;
            if (Math.abs(rowX - x) < epsilon && Math.abs(rowY - y) < epsilon) {
                // Apply main filter if present
                if (useFilter && mainFilterValue !== '') {
                    if (!valueMatches(row[mainFilterColumnIndex], mainFilterValue)) {
                        continue;
                    }
                }
                
                matchingRows.push(row);
                break; // No need to check other y columns for this row
            }
        }
    }
    
    return matchingRows;
}

// Update chart with new data (for filtering)
function updateChartWithData(newData) {
    if (!chartInstance) return;
    
    // For percentage stacked bar chart, need to recalculate percentages
    if (selectedChartType === 'percentStackedBar') {
        const processedData = {
            labels: newData.labels,
            datasets: calculatePercentageData(newData.datasets, newData.labels)
        };
        
        chartInstance.data.labels = processedData.labels;
        chartInstance.data.datasets.forEach((dataset, i) => {
            if (i < processedData.datasets.length) {
                dataset.data = processedData.datasets[i].data;
            }
        });
    } else {
        // For other chart types
        chartInstance.data.labels = newData.labels;
        chartInstance.data.datasets.forEach((dataset, i) => {
            if (i < newData.datasets.length) {
                dataset.data = newData.datasets[i].data;
            }
        });
    }
    
    chartInstance.update();
}