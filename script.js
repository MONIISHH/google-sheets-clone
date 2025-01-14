document.addEventListener('DOMContentLoaded', () => {
    // ========== PART A: Create the Spreadsheet ==========
    const spreadsheet = document.getElementById('spreadsheet');
    const formulaInput = document.getElementById('formula-input');
    let rows = 40;
    let cols = 26;
    let isDragging = false;
    let startCell = null;
    const cellDependencies = {};

    function createSpreadsheet(rows, cols) {
        spreadsheet.innerHTML = '';
        for (let i = 0; i <= rows; i++) {
            const row = document.createElement('tr');
            for (let j = 0; j <= cols; j++) {
                const cell = document.createElement(i === 0 || j === 0 ? 'th' : 'td');
                if (i === 0 && j > 0) {
                    cell.textContent = getColumnName(j); // A, B, C...
                } else if (j === 0 && i > 0) {
                    cell.textContent = i; // 1, 2, 3...
                } else if (i > 0 && j > 0) {
                    cell.contentEditable = true;
                    cell.dataset.row = i;
                    cell.dataset.col = j;
                    cell.dataset.type = 'text'; // Allow text input
                    cell.addEventListener('input', () => updateDependencies(cell));
                    cell.addEventListener('blur', () => validateCell(cell));
                }
                row.appendChild(cell);
            }
            spreadsheet.appendChild(row);
        }
    }

    function getColumnName(index) {
        let columnName = '';
        while (index > 0) {
            const remainder = (index - 1) % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            index = Math.floor((index - 1) / 26);
        }
        return columnName;
    }

    function validateCell(cell) {
        const value = cell.textContent;
        if (cell.dataset.type === 'number' && isNaN(value)) {
            alert('Please enter a valid number.');
            cell.textContent = '';
        } else if (cell.dataset.type === 'date' && isNaN(Date.parse(value))) {
            alert('Please enter a valid date.');
            cell.textContent = '';
        }
        // Remove the alert for valid text
        // else if (cell.dataset.type === 'text' && value.trim() === '') {
        //     alert('Please enter valid text.');
        //     cell.textContent = '';
        // }
    }

    createSpreadsheet(rows, cols);

    // ========== PART B: Add Mathematical Functions ==========
    function sum(range) {
        const cells = getRangeCells(range);
        return cells.reduce((acc, cell) => acc + parseFloat(cell.textContent || 0), 0);
    }

    function average(range) {
        const cells = getRangeCells(range);
        const validCells = cells.map(cell => parseFloat(cell.textContent || 0)).filter(val => !isNaN(val));
        return validCells.reduce((acc, val) => acc + val, 0) / validCells.length;
    }

    function max(range) {
        const cells = getRangeCells(range);
        const validCells = cells.map(cell => parseFloat(cell.textContent || 0)).filter(val => !isNaN(val));
        return Math.max(...validCells);
    }

    function min(range) {
        const cells = getRangeCells(range);
        const validCells = cells.map(cell => parseFloat(cell.textContent || 0)).filter(val => !isNaN(val));
        return Math.min(...validCells);
    }

    function count(range) {
        const cells = getRangeCells(range);
        return cells.filter(cell => !isNaN(parseFloat(cell.textContent))).length;
    }

    function median(range) {
        const cells = getRangeCells(range);
        const validCells = cells.map(cell => parseFloat(cell.textContent || 0)).filter(val => !isNaN(val)).sort((a, b) => a - b);
        const mid = Math.floor(validCells.length / 2);
        return validCells.length % 2 !== 0 ? validCells[mid] : (validCells[mid - 1] + validCells[mid]) / 2;
    }

    function mode(range) {
        const cells = getRangeCells(range);
        const validCells = cells.map(cell => parseFloat(cell.textContent || 0)).filter(val => !isNaN(val));
        const frequency = {};
        validCells.forEach(val => frequency[val] = (frequency[val] || 0) + 1);
        const maxFreq = Math.max(...Object.values(frequency));
        return parseFloat(Object.keys(frequency).find(key => frequency[key] === maxFreq));
    }

    // Data Quality Functions
    function TRIM(cell) {
        return cell.trim();
    }

    function UPPER(cell) {
        return cell.toUpperCase();
    }

    function LOWER(cell) {
        return cell.toLowerCase();
    }

    function REMOVE_DUPLICATES(range) {
        return range.filter((row, index, self) =>
            index === self.findIndex((r) => JSON.stringify(r) === JSON.stringify(row))
        );
    }

    function FIND_AND_REPLACE(range, findText, replaceText) {
        return range.map(row =>
            row.map(cell => cell.replace(new RegExp(findText, 'g'), replaceText))
        );
    }

    // Add buttons for data quality functions
    document.getElementById('trim-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            const result = TRIM(cell.textContent);
            cell.textContent = result;
            alert(`The trimmed value is "${result}"`);
        } else {
            alert('Please select a cell first.');
        }
    });

    document.getElementById('upper-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            const result = UPPER(cell.textContent);
            cell.textContent = result;
            alert(`The uppercase value is "${result}"`);
        } else {
            alert('Please select a cell first.');
        }
    });

    document.getElementById('lower-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            const result = LOWER(cell.textContent);
            cell.textContent = result;
            alert(`The lowercase value is "${result}"`);
        } else {
            alert('Please select a cell first.');
        }
    });

    document.getElementById('remove-duplicates-btn').addEventListener('click', () => {
        const range = getRangeCells('A1:Z40'); // Replace with your range
        const result = REMOVE_DUPLICATES(range.map(cell => cell.textContent));
        result.forEach((text, index) => {
            range[index].textContent = text;
        });
        alert(`The range without duplicates is ${JSON.stringify(result)}`);
    });

    document.getElementById('find-replace-btn').addEventListener('click', () => {
        const range = getRangeCells('A1:Z40'); // Replace with your range
        const findText = prompt('Enter the text to find:');
        let found = false;
        range.forEach(cell => {
            if (cell.textContent.includes(findText)) {
                found = true;
            }
        });
        if (found) {
            const replaceText = prompt('Enter the text to replace with:');
            range.forEach(cell => {
                cell.textContent = cell.textContent.replace(new RegExp(findText, 'g'), replaceText);
            });
            alert('Text replaced successfully.');
        } else {
            alert('Text not found.');
        }
    });

    // Add buttons for text formatting
    document.getElementById('bold-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            cell.style.fontWeight = cell.style.fontWeight === 'bold' ? 'normal' : 'bold';
        } else {
            alert('Please select a cell first.');
        }
    });

    document.getElementById('italic-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            cell.style.fontStyle = cell.style.fontStyle === 'italic' ? 'normal' : 'italic';
        } else {
            alert('Please select a cell first.');
        }
    });

    document.getElementById('underline-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            cell.style.textDecoration = cell.style.textDecoration === 'underline' ? 'none' : 'underline';
        } else {
            alert('Please select a cell first.');
        }
    });

    document.getElementById('font-size-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            const fontSize = prompt('Enter the font size (e.g., 12px, 1em):');
            cell.style.fontSize = fontSize;
        } else {
            alert('Please select a cell first.');
        }
    });

    document.getElementById('font-color-btn').addEventListener('click', () => {
        const cell = document.querySelector('td.active');
        if (cell) {
            const fontColor = prompt('Enter the font color (e.g., red, #ff0000):');
            cell.style.color = fontColor;
        } else {
            alert('Please select a cell first.');
        }
    });

    // Add event listener for font family dropdown
    document.getElementById('font-family-dropdown').addEventListener('change', (event) => {
        const cell = document.querySelector('td.active');
        if (cell) {
            cell.style.fontFamily = event.target.value;
        } else {
            alert('Please select a cell first.');
        }
    });

    // Add buttons for row and column operations
    document.getElementById('add-row-btn').addEventListener('click', () => {
        const newRow = document.createElement('tr');
        for (let j = 0; j <= cols; j++) {
            const cell = document.createElement(j === 0 ? 'th' : 'td');
            if (j === 0) {
                cell.textContent = rows + 1;
            } else {
                cell.contentEditable = true;
                cell.dataset.row = rows + 1;
                cell.dataset.col = j;
                cell.addEventListener('input', () => updateDependencies(cell));
                cell.addEventListener('blur', () => validateCell(cell));
            }
            newRow.appendChild(cell);
        }
        spreadsheet.appendChild(newRow);
        rows++;
    });

    document.getElementById('delete-row-btn').addEventListener('click', () => {
        if (rows > 1) {
            spreadsheet.deleteRow(rows);
            rows--;
        } else {
            alert('Cannot delete the last row.');
        }
    });

    document.getElementById('add-column-btn').addEventListener('click', () => {
        const rows = spreadsheet.rows;
        for (let i = 0; i < rows.length; i++) {
            const cell = document.createElement(i === 0 ? 'th' : 'td');
            if (i === 0) {
                cell.textContent = getColumnName(cols + 1);
            } else {
                cell.contentEditable = true;
                cell.dataset.row = i;
                cell.dataset.col = cols + 1;
                cell.addEventListener('input', () => updateDependencies(cell));
                cell.addEventListener('blur', () => validateCell(cell));
            }
            rows[i].appendChild(cell);
        }
        cols++;
    });

    document.getElementById('delete-column-btn').addEventListener('click', () => {
        if (cols > 1) {
            const rows = spreadsheet.rows;
            for (let i = 0; i < rows.length; i++) {
                rows[i].deleteCell(cols);
            }
            cols--;
        } else {
            alert('Cannot delete the last column.');
        }
    });

    document.getElementById('resize-row-btn').addEventListener('click', () => {
        const rowIndex = prompt('Enter the row number to resize:');
        const newSize = prompt('Enter the new height (e.g., 30px):');
        const row = spreadsheet.rows[rowIndex];
        if (row) {
            for (let cell of row.cells) {
                cell.style.height = newSize;
            }
        } else {
            alert('Invalid row number.');
        }
    });

    document.getElementById('resize-column-btn').addEventListener('click', () => {
        const colIndex = prompt('Enter the column letter to resize:').toUpperCase().charCodeAt(0) - 64;
        const newSize = prompt('Enter the new width (e.g., 100px):');
        for (let i = 0; i <= rows; i++) {
            const cell = document.querySelector(`td[data-row="${i}"][data-col="${colIndex}"]`) || document.querySelector(`th[data-col="${colIndex}"]`);
            if (cell) {
                cell.style.width = newSize;
            }
        }
    });

    // Save and Load functions using localStorage
    document.getElementById('save-btn').addEventListener('click', () => {
        const data = [];
        for (let i = 1; i <= rows; i++) {
            const row = [];
            for (let j = 1; j <= cols; j++) {
                const cell = document.querySelector(`td[data-row="${i}"][data-col="${j}"]`);
                row.push({
                    text: cell.textContent,
                    style: cell.getAttribute('style')
                });
            }
            data.push(row);
        }
        localStorage.setItem('spreadsheetData', JSON.stringify(data));
        alert('Spreadsheet saved successfully.');
    });

    document.getElementById('load-btn').addEventListener('click', () => {
        const data = JSON.parse(localStorage.getItem('spreadsheetData'));
        if (data) {
            createSpreadsheet(data.length, data[0].length);
            for (let i = 1; i <= data.length; i++) {
                for (let j = 1; j <= data[0].length; j++) {
                    const cell = document.querySelector(`td[data-row="${i}"][data-col="${j}"]`);
                    cell.textContent = data[i - 1][j - 1].text || '';
                    cell.setAttribute('style', data[i - 1][j - 1].style || '');
                }
            }
            alert('Spreadsheet loaded successfully.');
        } else {
            alert('No saved spreadsheet found.');
        }
    });

    function getRangeCells(range) {
        const [start, end] = range.split(':');
        const startRow = parseInt(start.slice(1), 10);
        const startCol = start.charCodeAt(0) - 65 + 1;
        const endRow = parseInt(end.slice(1), 10);
        const endCol = end.charCodeAt(0) - 65 + 1;

        const cells = [];
        for (let i = startRow; i <= endRow; i++) {
            for (let j = startCol; j <= endCol; j++) {
                const cell = document.querySelector(`td[data-row="${i}"][data-col="${j}"]`);
                if (cell) {
                    cells.push(cell);
                }
            }
        }
        return cells;
    }

    // Example SUM button
    document.getElementById('sum-btn').addEventListener('click', () => {
        const result = sum('A1:Z40'); // Replace with your range
        alert(`The sum is ${result}`);
    });

    // AVERAGE Button
    document.getElementById('average-btn').addEventListener('click', () => {
        const result = average('A1:Z40'); // Replace with your range
        alert(`The average is ${result}`);
    });

    // MAX Button
    document.getElementById('max-btn').addEventListener('click', () => {
        const result = max('A1:Z40'); // Replace with your range
        alert(`The max value is ${result}`);
    });

    // MIN Button
    document.getElementById('min-btn').addEventListener('click', () => {
        const result = min('A1:Z40'); // Replace with your range
        alert(`The min value is ${result}`);
    });

    // COUNT Button
    document.getElementById('count-btn').addEventListener('click', () => {
        const result = count('A1:Z40'); // Replace with your range
        alert(`The count is ${result}`);
    });

    // MEDIAN Button
    document.getElementById('median-btn').addEventListener('click', () => {
        const result = median('A1:Z40'); // Replace with your range
        alert(`The median is ${result}`);
    });

    // MODE Button
    document.getElementById('mode-btn').addEventListener('click', () => {
        const result = mode('A1:Z40'); // Replace with your range
        alert(`The mode is ${result}`);
    });

    // Add event listeners for test buttons
    document.getElementById('testSumBtn').addEventListener('click', () => {
        const range = document.getElementById('testRange').value;
        const result = sum(range);
        document.getElementById('resultText').textContent = `The sum of ${range} is ${result}`;
    });

    document.getElementById('testAverageBtn').addEventListener('click', () => {
        const range = document.getElementById('testRange').value;
        const result = average(range);
        document.getElementById('resultText').textContent = `The average of ${range} is ${result}`;
    });

    document.getElementById('testMaxBtn').addEventListener('click', () => {
        const range = document.getElementById('testRange').value;
        const result = max(range);
        document.getElementById('resultText').textContent = `The max value of ${range} is ${result}`;
    });

    document.getElementById('testMinBtn').addEventListener('click', () => {
        const range = document.getElementById('testRange').value;
        const result = min(range);
        document.getElementById('resultText').textContent = `The min value of ${range} is ${result}`;
    });

    document.getElementById('testCountBtn').addEventListener('click', () => {
        const range = document.getElementById('testRange').value;
        const result = count(range);
        document.getElementById('resultText').textContent = `The count of ${range} is ${result}`;
    });

    document.getElementById('testMedianBtn').addEventListener('click', () => {
        const range = document.getElementById('testRange').value;
        const result = median(range);
        document.getElementById('resultText').textContent = `The median of ${range} is ${result}`;
    });

    document.getElementById('testModeBtn').addEventListener('click', () => {
        const range = document.getElementById('testRange').value;
        const result = mode(range);
        document.getElementById('resultText').textContent = `The mode of ${range} is ${result}`;
    });

    // ========== PART D: Add Selection and Styling ==========
    document.addEventListener('mousedown', (event) => {
        if (event.target.tagName === 'TD') {
            isDragging = true;
            startCell = event.target;
            document.querySelectorAll('td.active').forEach(cell => cell.classList.remove('active'));
            event.target.classList.add('active');
            formulaInput.value = event.target.textContent;
        }
    });

    document.addEventListener('mousemove', (event) => {
        if (isDragging && event.target.tagName === 'TD') {
            const endCell = event.target;
            const startRow = parseInt(startCell.dataset.row, 10);
            const startCol = parseInt(startCell.dataset.col, 10);
            const endRow = parseInt(endCell.dataset.row, 10);
            const endCol = parseInt(endCell.dataset.col, 10);

            document.querySelectorAll('td').forEach(cell => cell.classList.remove('selected'));
            for (let i = Math.min(startRow, endRow); i <= Math.max(startRow, endRow); i++) {
                for (let j = Math.min(startCol, endCol); j <= Math.max(startCol, endCol); j++) {
                    document.querySelector(`td[data-row="${i}"][data-col="${j}"]`).classList.add('selected');
                }
            }
        }
    });

    document.addEventListener('mouseup', () => {
        isDragging = false;
    });

    formulaInput.addEventListener('input', () => {
        const activeCell = document.querySelector('td.active');
        if (activeCell) {
            activeCell.textContent = formulaInput.value;
        }
    });

    formulaInput.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
            const activeCell = document.querySelector('td.active');
            if (activeCell) {
                const formula = formulaInput.value;
                if (formula.startsWith('=')) {
                    try {
                        const result = evaluateFormula(formula.slice(1), activeCell);
                        activeCell.textContent = result;
                        updateDependencies(activeCell);
                    } catch (error) {
                        alert('Error in formula');
                    }
                } else {
                    activeCell.textContent = formula;
                }
            }
        }
    });

    function evaluateFormula(formula, activeCell) {
        const cellReferenceRegex = /([A-Z]+)(\d+)/g;
        const evaluatedFormula = formula.replace(cellReferenceRegex, (match, col, row) => {
            const colIndex = col.charCodeAt(0) - 65 + 1;
            const rowIndex = parseInt(row, 10);
            const cell = document.querySelector(`td[data-row="${rowIndex}"][data-col="${colIndex}"]`);
            addDependency(cell, activeCell);
            return cell ? parseFloat(cell.textContent || 0) : 0;
        });
        return new Function('return ' + evaluatedFormula)();
    }

    function addDependency(sourceCell, dependentCell) {
        const sourceKey = `${sourceCell.dataset.row},${sourceCell.dataset.col}`;
        if (!cellDependencies[sourceKey]) {
            cellDependencies[sourceKey] = [];
        }
        cellDependencies[sourceKey].push(dependentCell);
    }

    function updateDependencies(cell) {
        const cellKey = `${cell.dataset.row},${cell.dataset.col}`;
        const dependents = cellDependencies[cellKey] || [];
        dependents.forEach(dependentCell => {
            const formula = dependentCell.textContent;
            if (formula.startsWith('=')) {
                try {
                    const result = evaluateFormula(formula.slice(1), dependentCell);
                    dependentCell.textContent = result;
                } catch (error) {
                    alert('Error in formula');
                }
            }
        });
    }

    // Load Chart.js library
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/chart.js';
    document.head.appendChild(script);

    script.onload = () => {
        function createChart(type, labels, data) {
            const ctx = document.getElementById('chart').getContext('2d');
            new Chart(ctx, {
                type: type,
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Dataset',
                        data: data,
                        backgroundColor: 'rgba(75, 192, 192, 0.2)',
                        borderColor: 'rgba(75, 192, 192, 1)',
                        borderWidth: 1
                    }]
                },
                options: {
                    scales: {
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
        }

        function getChartData(range) {
            const cells = getRangeCells(range);
            const labels = cells.map(cell => cell.dataset.row + cell.dataset.col);
            const data = cells.map(cell => parseFloat(cell.textContent) || 0);
            return { labels, data };
        }

        document.getElementById('bar-chart-btn').addEventListener('click', () => {
            const range = prompt('Enter the range for the chart (e.g., A1:B2):');
            if (range) {
                const { labels, data } = getChartData(range);
                createChart('bar', labels, data);
            }
        });

        document.getElementById('line-chart-btn').addEventListener('click', () => {
            const range = prompt('Enter the range for the chart (e.g., A1:B2):');
            if (range) {
                const { labels, data } = getChartData(range);
                createChart('line', labels, data);
            }
        });

        document.getElementById('pie-chart-btn').addEventListener('click', () => {
            const range = prompt('Enter the range for the chart (e.g., A1:B2):');
            if (range) {
                const { labels, data } = getChartData(range);
                createChart('pie', labels, data);
            }
        });
    };
});

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('dataEntryForm');
    
    form.addEventListener('submit', (event) => {
        event.preventDefault();
        
        const numberInput = document.getElementById('numberInput').value;
        const textInput = document.getElementById('textInput').value;
        const dateInput = document.getElementById('dateInput').value;
        
        if (isNaN(numberInput) || numberInput.trim() === '') {
            alert('Please enter a valid number.');
            return;
        }
        
        // Additional validation checks can be added here
        
        alert('Form submitted successfully!');
    });
});

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('dataEntryForm');
    
    form.addEventListener('submit', (event) => {
        event.preventDefault();
        
        const numberInput = document.getElementById('numberInput').value;
        const textInput = document.getElementById('textInput').value;
        const dateInput = document.getElementById('dateInput').value;
        
        if (isNaN(numberInput) || numberInput.trim() === '') {
            alert('Please enter a valid number.');
            return;
        }
        
        // Additional validation checks can be added here
        
        alert('Form submitted successfully!');
    });
});
