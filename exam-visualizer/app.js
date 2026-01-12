// ===== Global Variables =====
let studentsData = [];
let charts = {};
let sortAscending = true;

// ===== Grade Colors =====
const gradeColors = {
    'A+': '#15803d',
    'A': '#059669',
    'A-': '#0d9488',
    'B+': '#2563eb',
    'B': '#4f46e5',
    'B-': '#7c3aed',
    'C+': '#d97706',
    'C': '#ea580c',
    'C-': '#c2410c',
    'D': '#dc2626',
    'F': '#b91c1c'
};

const gradeBackgrounds = {
    'A+': '#dcfce7',
    'A': '#d1fae5',
    'A-': '#ccfbf1',
    'B+': '#dbeafe',
    'B': '#e0e7ff',
    'B-': '#ede9fe',
    'C+': '#fef3c7',
    'C': '#fed7aa',
    'C-': '#ffedd5',
    'D': '#fee2e2',
    'F': '#fecaca'
};

// ===== DOM Elements =====
const uploadArea = document.getElementById('upload-area');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const uploadSection = document.getElementById('upload-section');
const dashboardSection = document.getElementById('dashboard-section');
const searchInput = document.getElementById('search-input');
const sortBtn = document.getElementById('sort-btn');
const printReportBtn = document.getElementById('print-report-btn');
const newFileBtn = document.getElementById('new-file-btn');

// ===== Event Listeners =====
document.addEventListener('DOMContentLoaded', () => {
    setupDragAndDrop();
    setupEventListeners();
});

function setupDragAndDrop() {
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    uploadArea.addEventListener('click', (e) => {
        if (e.target.tagName !== 'BUTTON') {
            fileInput.click();
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });
}

function setupEventListeners() {
    searchInput.addEventListener('input', filterTable);
    sortBtn.addEventListener('click', toggleSort);
    printReportBtn.addEventListener('click', printReport);
    newFileBtn.addEventListener('click', resetToUpload);
}

// ===== File Handling =====
function handleFile(file) {
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    const fileName = file.name.toLowerCase();

    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
        showError('Please upload a valid Excel file (.xlsx or .xls)');
        return;
    }

    fileInfo.innerHTML = `<i class="fas fa-check-circle"></i> Loading: ${file.name}`;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);

            if (jsonData.length === 0) {
                showError('The Excel file is empty or has no valid data');
                return;
            }

            processData(jsonData, file.name);
        } catch (error) {
            showError('Error reading Excel file: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

function showError(message) {
    fileInfo.innerHTML = `<span style="color: #dc3545;"><i class="fas fa-exclamation-circle"></i> ${message}</span>`;
}

// ===== Data Processing =====
function processData(data, fileName) {
    // Normalize column names
    studentsData = data.map((row, index) => ({
        id: index + 1,
        student: row['Student'] || row['Name'] || row['Student Name'] || '',
        nic: row['NIC/Passport'] || row['NIC'] || row['Passport'] || '',
        registrationNo: row['Registration No'] || row['Reg No'] || row['Registration Number'] || '',
        admissionNo: row['Admission No'] || row['Admission Number'] || '',
        subjectMarks: parseFloat(row['Subject Marks'] || row['Exam Marks'] || 0),
        assessmentMarks: parseFloat(row['Assessment Marks'] || row['CA Marks'] || 0),
        finalMarks: parseInt(row['Final Marks'] || row['Total Marks'] || row['Total'] || 0),
        grade: row['Grade'] || '',
        absentCount: parseInt(row['Absent Count'] || 0),
        status: row['Status'] || 'Unknown',
        remark: row['Remark'] || ''
    }));

    // Extract batch name from registration number or filename
    let batchName = 'Exam Results';
    if (studentsData.length > 0 && studentsData[0].registrationNo) {
        const regNo = studentsData[0].registrationNo;
        const match = regNo.match(/^([A-Z]+\/\d{4}-\d+[A-Z]*\/[A-Z]+)/);
        if (match) {
            batchName = match[1];
        }
    }
    if (batchName === 'Exam Results') {
        batchName = fileName.replace('.xlsx', '').replace('.xls', '');
    }

    fileInfo.innerHTML = `<i class="fas fa-check-circle"></i> Successfully loaded: ${fileName} (${studentsData.length} students)`;

    setTimeout(() => {
        uploadSection.classList.add('hidden');
        dashboardSection.classList.remove('hidden');
        renderDashboard(batchName);
    }, 500);
}

// ===== Dashboard Rendering =====
function renderDashboard(batchName) {
    document.getElementById('batch-name').textContent = batchName;

    // Calculate statistics
    const stats = calculateStats();

    // Update stat cards
    document.getElementById('total-students').textContent = stats.totalStudents;
    document.getElementById('avg-marks').textContent = stats.avgMarks.toFixed(1);
    document.getElementById('highest-marks').textContent = `${stats.highestMarks} (${stats.highestGrade})`;
    document.getElementById('lowest-marks').textContent = `${stats.lowestMarks} (${stats.lowestGrade})`;
    document.getElementById('avg-subject').textContent = stats.avgSubject.toFixed(1);
    document.getElementById('avg-assessment').textContent = stats.avgAssessment.toFixed(1);
    document.getElementById('pass-rate').textContent = `${stats.passRate.toFixed(1)}%`;

    // Render components
    renderGradeSummary(stats.gradeDistribution);
    renderCharts(stats);
    renderTable();
}

function calculateStats() {
    const totalStudents = studentsData.length;
    const finalMarks = studentsData.map(s => s.finalMarks);
    const subjectMarks = studentsData.map(s => s.subjectMarks);
    const assessmentMarks = studentsData.map(s => s.assessmentMarks);

    const avgMarks = finalMarks.reduce((a, b) => a + b, 0) / totalStudents;
    const avgSubject = subjectMarks.reduce((a, b) => a + b, 0) / totalStudents;
    const avgAssessment = assessmentMarks.reduce((a, b) => a + b, 0) / totalStudents;

    const highestMarks = Math.max(...finalMarks);
    const lowestMarks = Math.min(...finalMarks);
    const highestStudent = studentsData.find(s => s.finalMarks === highestMarks);
    const lowestStudent = studentsData.find(s => s.finalMarks === lowestMarks);

    // Grade distribution
    const gradeDistribution = {};
    studentsData.forEach(s => {
        const grade = s.grade || 'N/A';
        gradeDistribution[grade] = (gradeDistribution[grade] || 0) + 1;
    });

    // Sort grades in order
    const gradeOrder = ['A+', 'A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D', 'F'];
    const sortedGradeDistribution = {};
    gradeOrder.forEach(grade => {
        if (gradeDistribution[grade]) {
            sortedGradeDistribution[grade] = gradeDistribution[grade];
        }
    });
    // Add any other grades not in the standard order
    Object.keys(gradeDistribution).forEach(grade => {
        if (!sortedGradeDistribution[grade]) {
            sortedGradeDistribution[grade] = gradeDistribution[grade];
        }
    });

    // Calculate pass rate (assuming C or above is pass)
    const passingGrades = ['A+', 'A', 'A-', 'B+', 'B', 'B-', 'C+', 'C'];
    const passCount = studentsData.filter(s => passingGrades.includes(s.grade)).length;
    const passRate = (passCount / totalStudents) * 100;

    // Marks distribution for histogram
    const marksRanges = {
        '90-100': 0,
        '80-89': 0,
        '70-79': 0,
        '60-69': 0,
        '50-59': 0,
        '40-49': 0,
        'Below 40': 0
    };

    studentsData.forEach(s => {
        const marks = s.finalMarks;
        if (marks >= 90) marksRanges['90-100']++;
        else if (marks >= 80) marksRanges['80-89']++;
        else if (marks >= 70) marksRanges['70-79']++;
        else if (marks >= 60) marksRanges['60-69']++;
        else if (marks >= 50) marksRanges['50-59']++;
        else if (marks >= 40) marksRanges['40-49']++;
        else marksRanges['Below 40']++;
    });

    return {
        totalStudents,
        avgMarks,
        avgSubject,
        avgAssessment,
        highestMarks,
        lowestMarks,
        highestGrade: highestStudent?.grade || 'N/A',
        lowestGrade: lowestStudent?.grade || 'N/A',
        gradeDistribution: sortedGradeDistribution,
        marksRanges,
        passRate
    };
}

// ===== Grade Summary =====
function renderGradeSummary(gradeDistribution) {
    const container = document.getElementById('grade-summary');
    const total = studentsData.length;

    container.innerHTML = Object.entries(gradeDistribution).map(([grade, count]) => {
        const percent = ((count / total) * 100).toFixed(1);
        const color = gradeColors[grade] || '#6c757d';
        const bg = gradeBackgrounds[grade] || '#f4f6f9';

        return `
            <div class="grade-item" style="background: ${bg};">
                <div class="grade-badge" style="color: ${color};">${grade}</div>
                <div class="grade-count">${count}</div>
                <div class="grade-percent">${percent}%</div>
            </div>
        `;
    }).join('');
}

// ===== Charts =====
function renderCharts(stats) {
    // Destroy existing charts
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};

    // Grade Distribution Doughnut Chart
    const gradeLabels = Object.keys(stats.gradeDistribution);
    const gradeData = Object.values(stats.gradeDistribution);
    const gradeChartColors = gradeLabels.map(g => gradeColors[g] || '#6c757d');

    charts.grade = new Chart(document.getElementById('grade-chart'), {
        type: 'doughnut',
        data: {
            labels: gradeLabels,
            datasets: [{
                data: gradeData,
                backgroundColor: gradeChartColors,
                borderWidth: 2,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        padding: 15,
                        usePointStyle: true,
                        font: { size: 12 }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function (context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percent = ((context.raw / total) * 100).toFixed(1);
                            return `${context.label}: ${context.raw} students (${percent}%)`;
                        }
                    }
                }
            }
        }
    });

    // Marks Distribution Bar Chart
    const marksLabels = Object.keys(stats.marksRanges);
    const marksData = Object.values(stats.marksRanges);

    charts.marks = new Chart(document.getElementById('marks-chart'), {
        type: 'bar',
        data: {
            labels: marksLabels,
            datasets: [{
                label: 'Number of Students',
                data: marksData,
                backgroundColor: [
                    '#15803d', '#059669', '#0d9488', '#2563eb', '#7c3aed', '#d97706', '#dc2626'
                ],
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { stepSize: 1 },
                    grid: { color: '#e2e8f0' }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });

    // Subject vs Assessment Comparison Chart
    const avgSubjectByGrade = {};
    const avgAssessmentByGrade = {};

    Object.keys(stats.gradeDistribution).forEach(grade => {
        const gradeStudents = studentsData.filter(s => s.grade === grade);
        avgSubjectByGrade[grade] = gradeStudents.reduce((a, b) => a + b.subjectMarks, 0) / gradeStudents.length;
        avgAssessmentByGrade[grade] = gradeStudents.reduce((a, b) => a + b.assessmentMarks, 0) / gradeStudents.length;
    });

    charts.comparison = new Chart(document.getElementById('comparison-chart'), {
        type: 'bar',
        data: {
            labels: Object.keys(stats.gradeDistribution),
            datasets: [
                {
                    label: 'Avg Subject Marks',
                    data: Object.values(avgSubjectByGrade),
                    backgroundColor: '#4361ee',
                    borderRadius: 6
                },
                {
                    label: 'Avg Assessment Marks',
                    data: Object.values(avgAssessmentByGrade),
                    backgroundColor: '#7c3aed',
                    borderRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: { usePointStyle: true }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: '#e2e8f0' }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });

    // Top 5 Performers Chart
    const topPerformers = [...studentsData]
        .sort((a, b) => b.finalMarks - a.finalMarks)
        .slice(0, 5);

    charts.topPerformers = new Chart(document.getElementById('top-performers-chart'), {
        type: 'bar',
        data: {
            labels: topPerformers.map(s => s.student.split(' ').slice(0, 2).join(' ')),
            datasets: [{
                label: 'Final Marks',
                data: topPerformers.map(s => s.finalMarks),
                backgroundColor: topPerformers.map((s, i) => {
                    const colors = ['#f59e0b', '#6b7280', '#c2410c', '#4361ee', '#7c3aed'];
                    return colors[i];
                }),
                borderRadius: 8
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    max: 100,
                    grid: { color: '#e2e8f0' }
                },
                y: {
                    grid: { display: false }
                }
            }
        }
    });
}

// ===== Table =====
function renderTable(data = studentsData) {
    const tbody = document.getElementById('table-body');
    const highestMarks = Math.max(...studentsData.map(s => s.finalMarks));
    const lowestMarks = Math.min(...studentsData.map(s => s.finalMarks));

    tbody.innerHTML = data.map((student, index) => {
        const gradeClass = getGradeClass(student.grade);
        const statusClass = student.status.toLowerCase() === 'processed' ? 'status-processed' : 'status-pending';

        let rowClass = '';
        if (student.finalMarks === highestMarks) rowClass = 'top-performer';
        else if (student.finalMarks === lowestMarks) rowClass = 'low-performer';

        return `
            <tr class="${rowClass}">
                <td>${index + 1}</td>
                <td>${student.student}</td>
                <td>${student.registrationNo}</td>
                <td>${student.subjectMarks.toFixed(1)}</td>
                <td>${student.assessmentMarks.toFixed(1)}</td>
                <td><strong>${student.finalMarks}</strong></td>
                <td><span class="grade-badge ${gradeClass}">${student.grade}</span></td>
                <td><span class="status-badge ${statusClass}">${student.status}</span></td>
            </tr>
        `;
    }).join('');
}

function getGradeClass(grade) {
    const gradeMap = {
        'A+': 'grade-A-plus',
        'A': 'grade-A',
        'A-': 'grade-A-minus',
        'B+': 'grade-B-plus',
        'B': 'grade-B',
        'B-': 'grade-B-minus',
        'C+': 'grade-C-plus',
        'C': 'grade-C',
        'C-': 'grade-C-minus',
        'D': 'grade-D',
        'F': 'grade-F'
    };
    return gradeMap[grade] || 'grade-C';
}

function filterTable() {
    const searchTerm = searchInput.value.toLowerCase();
    const filtered = studentsData.filter(s =>
        s.student.toLowerCase().includes(searchTerm) ||
        s.registrationNo.toLowerCase().includes(searchTerm) ||
        s.grade.toLowerCase().includes(searchTerm)
    );
    renderTable(filtered);
}

function toggleSort() {
    sortAscending = !sortAscending;
    const sorted = [...studentsData].sort((a, b) =>
        sortAscending ? a.finalMarks - b.finalMarks : b.finalMarks - a.finalMarks
    );
    renderTable(sorted);
    sortBtn.innerHTML = sortAscending
        ? '<i class="fas fa-sort-amount-up"></i> Sort Descending'
        : '<i class="fas fa-sort-amount-down"></i> Sort Ascending';
}

// ===== Print Report =====
function printReport() {
    const stats = calculateStats();
    const batchName = document.getElementById('batch-name').textContent;
    const today = new Date().toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric'
    });

    // Populate print section
    document.getElementById('print-batch-name').textContent = batchName;
    document.getElementById('print-date').textContent = `Generated on: ${today}`;

    // Print stats
    document.getElementById('print-stats').innerHTML = `
        <div class="print-stat-item">
            <h4>${stats.totalStudents}</h4>
            <p>Total Students</p>
        </div>
        <div class="print-stat-item">
            <h4>${stats.avgMarks.toFixed(1)}</h4>
            <p>Average Marks</p>
        </div>
        <div class="print-stat-item">
            <h4>${stats.highestMarks}</h4>
            <p>Highest Marks</p>
        </div>
        <div class="print-stat-item">
            <h4>${stats.passRate.toFixed(1)}%</h4>
            <p>Pass Rate</p>
        </div>
    `;

    // Capture chart images for print - use Chart.js toBase64Image method
    const gradeChartImg = document.getElementById('print-grade-chart');
    const marksChartImg = document.getElementById('print-marks-chart');

    // Make sure charts exist and convert to images
    if (charts.grade && charts.marks) {
        gradeChartImg.src = charts.grade.toBase64Image('image/png', 1);
        marksChartImg.src = charts.marks.toBase64Image('image/png', 1);
    } else {
        // Fallback to canvas method
        const gradeChartCanvas = document.getElementById('grade-chart');
        const marksChartCanvas = document.getElementById('marks-chart');
        gradeChartImg.src = gradeChartCanvas.toDataURL('image/png', 1);
        marksChartImg.src = marksChartCanvas.toDataURL('image/png', 1);
    }

    // Print grade summary
    document.getElementById('print-grade-summary').innerHTML = `
        <h3>Grade Summary</h3>
        <div class="print-grade-grid">
            ${Object.entries(stats.gradeDistribution).map(([grade, count]) => {
        const percent = ((count / stats.totalStudents) * 100).toFixed(1);
        return `<div class="print-grade-item"><strong>${grade}</strong>: ${count} (${percent}%)</div>`;
    }).join('')}
        </div>
    `;

    // Print table
    document.getElementById('print-table').innerHTML = `
        <h3>Student Results</h3>
        <table>
            <thead>
                <tr>
                    <th>#</th>
                    <th>Student Name</th>
                    <th>Registration No</th>
                    <th>Subject</th>
                    <th>Assessment</th>
                    <th>Final</th>
                    <th>Grade</th>
                </tr>
            </thead>
            <tbody>
                ${[...studentsData].sort((a, b) => b.finalMarks - a.finalMarks).map((student, index) => `
                    <tr>
                        <td>${index + 1}</td>
                        <td>${student.student}</td>
                        <td>${student.registrationNo}</td>
                        <td>${student.subjectMarks.toFixed(1)}</td>
                        <td>${student.assessmentMarks.toFixed(1)}</td>
                        <td><strong>${student.finalMarks}</strong></td>
                        <td><span class="grade-badge ${getGradeClass(student.grade)}">${student.grade}</span></td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;

    // Wait for images to load before printing
    setTimeout(() => {
        window.print();
    }, 100);
}

// ===== Reset =====
function resetToUpload() {
    studentsData = [];
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};

    dashboardSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    fileInfo.innerHTML = '';
    fileInput.value = '';
    searchInput.value = '';
    sortAscending = true;
    sortBtn.innerHTML = '<i class="fas fa-sort"></i> Sort by Marks';
}
