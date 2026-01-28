import React, { useState } from 'react';
import { Upload, Download, FileSpreadsheet, X, Plus } from 'lucide-react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

const OCIMetricsUpdater = () => {
  const [csvFiles, setCsvFiles] = useState({
    cpuMean: null,
    cpuMax: null,
    memoryMean: null,
    memoryMax: null
  });
  const [logs, setLogs] = useState([]);
  const [compartmentName, setCompartmentName] = useState('');

  const addLog = (message, type = 'info') => {
    setLogs(prev => [...prev, { message, type, time: new Date().toLocaleTimeString() }]);
  };

  const handleCsvUpload = (type, e) => {
    const file = e.target.files[0];
    if (!file) return;

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: true,
      complete: (results) => {
        setCsvFiles(prev => ({ ...prev, [type]: results.data }));
        addLog(`${type.replace(/([A-Z])/g, ' $1').toUpperCase()} loaded: ${results.data.length} rows`, 'success');
      },
      error: (error) => {
        addLog(`Error parsing ${type} CSV: ` + error.message, 'error');
      }
    });
  };

  const clearCsv = (type) => {
    setCsvFiles(prev => ({ ...prev, [type]: null }));
    addLog(`${type.replace(/([A-Z])/g, ' $1').toUpperCase()} removed`, 'info');
  };

  const customOrder = [
    'HUBOVPNVM',
    'HUBADVM',
    'UATCACWTVM',
    'UATCACOLAVM',
    'UATCACPACIVM',
    'UATCACMWVM',
    'UATCACSHDVM',
    'UATWLSFREVM',
    'UATWLSBAEVM',
    'UATWLSLOGVM',
    'UATWRBFREVM',
    'UATWRBBAEVM',
    'UATANLKAFVM',
    'UATANLPBCNVM',
    'UATANLCLHDBVM',
    'PRDANLKAFVM1',
    'PRDANLKAFVM2',
    'PRDANLKAFVM3',
    'PRDANLPUBVM1',
    'PRDANLPUBVM2',
    'PRDANLCNRPVM1',
    'PRDANLCNRPVM2',
    'PRDCACTSNVM1',
    'PRDCACTSNVM2',
    'PRDCACOAUVM1',
    'PRDCACOAUVM2',
    'PRDCACFSVM1',
    'PRDCACFSVM2',
    'PRDCACSHDVM1',
    'PRDCACSHDVM2',
    'PRDCACADVM',
    'PRDCACWTVM',
    'PRDCACMWVM',
    'PRDWLSFREVM1',
    'PRDWLSFREVM2',
    'PRDWLSBAEVM1',
    'PRDWLSBAEVM2',
    'PRDWRBFREVM1',
    'PRDWRBFREVM2',
    'PRDWRBBAEVM1',
    'PRDWRBBAEVM2',
    'PRDANLCLHDBVM1',
    'PRDANLCLHDBVM2',
    'PRDWLSCLHDBVM1',
    'PRDWLSCLHDBVM2'
  ];

  const getAllInstanceNames = () => {
    const instanceNames = new Set();
    
    Object.values(csvFiles).forEach(csvData => {
      if (!csvData || csvData.length === 0) return;
      
      // Get all column names except 'group'
      const firstRow = csvData[0];
      Object.keys(firstRow).forEach(key => {
        if (key !== 'group' && key.trim() !== '') {
          instanceNames.add(key.trim());
        }
      });
    });
    
    const foundInstances = Array.from(instanceNames);
    
    // Sort by custom order
    return foundInstances.sort((a, b) => {
      const indexA = customOrder.indexOf(a);
      const indexB = customOrder.indexOf(b);
      
      // If both are in custom order, sort by their position
      if (indexA !== -1 && indexB !== -1) {
        return indexA - indexB;
      }
      
      // If only A is in custom order, A comes first
      if (indexA !== -1) return -1;
      
      // If only B is in custom order, B comes first
      if (indexB !== -1) return 1;
      
      // If neither is in custom order, sort alphabetically
      return a.localeCompare(b);
    });
  };

  const getLatestMetricValue = (csvData, instanceName) => {
    if (!csvData || csvData.length === 0) return null;
    
    // Find the column with this instance name
    const firstRow = csvData[0];
    const columnKey = Object.keys(firstRow).find(key => 
      key.trim() === instanceName
    );
    
    if (!columnKey) return null;
    
    // Get the latest (first) row's value for this instance
    const latestValue = csvData[0][columnKey];
    return typeof latestValue === 'number' ? latestValue : null;
  };

  const generateExcel = () => {
    const hasAnyCsv = Object.values(csvFiles).some(csv => csv !== null);
    if (!hasAnyCsv) {
      addLog('Please upload at least one CSV file', 'error');
      return;
    }

    try {
      addLog('Generating Excel file...', 'info');
      
      const instanceNames = getAllInstanceNames();
      
      if (instanceNames.length === 0) {
        addLog('No instances found in CSV files', 'error');
        return;
      }

      // Create worksheet data
      const wsData = [];
      
      // Add header row
      wsData.push([
        'Compartment',
        'Instance Display Name',
        'CPU Mean (%)',
        'CPU Max (%)',
        'Memory Mean (%)',
        'Memory Max (%)'
      ]);

      // Helper function to round and ensure minimum value of 1
      const formatValue = (value) => {
        if (value === null) return 'N/A';
        const rounded = Math.round(value);
        return rounded < 1 ? 1 : rounded;
      };

      // Add data rows for each instance
      instanceNames.forEach(instanceName => {
        const cpuMean = getLatestMetricValue(csvFiles.cpuMean, instanceName);
        const cpuMax = getLatestMetricValue(csvFiles.cpuMax, instanceName);
        const memoryMean = getLatestMetricValue(csvFiles.memoryMean, instanceName);
        const memoryMax = getLatestMetricValue(csvFiles.memoryMax, instanceName);

        wsData.push([
          compartmentName || 'N/A',
          instanceName,
          formatValue(cpuMean),
          formatValue(cpuMax),
          formatValue(memoryMean),
          formatValue(memoryMax)
        ]);

        addLog(`✓ Added: ${instanceName}`, 'success');
      });

      // Create workbook and worksheet
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(wsData);

      // Set column widths
      ws['!cols'] = [
        { wch: 20 }, // Compartment
        { wch: 25 }, // Instance Name
        { wch: 15 }, // CPU Mean
        { wch: 15 }, // CPU Max
        { wch: 15 }, // Memory Mean
        { wch: 15 }  // Memory Max
      ];

      // Format cells: Add % sign and center alignment
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let row = 1; row <= range.e.r; row++) { // Start from row 1 (skip header)
        for (let col = 2; col <= 5; col++) { // Columns C to F (CPU and Memory columns)
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = ws[cellAddress];
          if (cell && cell.v !== 'N/A') {
            // Add % sign to the value
            cell.t = 's'; // Set type to string
            cell.v = cell.v + '%';
            // Center alignment
            cell.s = {
              alignment: { horizontal: 'center', vertical: 'center' }
            };
          }
        }
      }

      // Center align header row
      for (let col = 0; col <= 5; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        const cell = ws[cellAddress];
        if (cell) {
          cell.s = {
            alignment: { horizontal: 'center', vertical: 'center' },
            font: { bold: true }
          };
        }
      }

      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(wb, ws, 'OCI Metrics');

      // Generate and download
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      const date = new Date().toISOString().split('T')[0];
      const compName = compartmentName ? `_${compartmentName.replace(/\s/g, '_')}` : '';
      link.download = `OCI_Metrics${compName}_${date}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);

      addLog(`Excel file generated! ${instanceNames.length} instances included`, 'success');
    } catch (error) {
      addLog('Error generating Excel: ' + error.message, 'error');
    }
  };

  const CsvUploadCard = ({ type, title, color }) => (
    <div className={`border-2 border-dashed rounded-lg p-4 transition-colors ${
      csvFiles[type] 
        ? `border-${color}-400 bg-${color}-50` 
        : 'border-gray-300 hover:border-gray-400'
    }`}>
      <label className="flex flex-col items-center cursor-pointer">
        <Upload className={`w-8 h-8 mb-2 ${
          csvFiles[type] ? `text-${color}-600` : 'text-gray-400'
        }`} />
        <span className="text-xs font-medium text-gray-700 text-center">{title}</span>
        <input
          type="file"
          accept=".csv"
          onChange={(e) => handleCsvUpload(type, e)}
          className="hidden"
        />
      </label>
      {csvFiles[type] && (
        <div className="mt-2 flex items-center justify-between">
          <span className="text-xs text-green-600 font-medium">
            ✓ {csvFiles[type].length} rows
          </span>
          <button
            onClick={() => clearCsv(type)}
            className="text-gray-500 hover:text-red-600 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-6">
      <div className="max-w-6xl mx-auto">
        <div className="bg-white rounded-lg shadow-xl p-8">
          <div className="flex items-center gap-3 mb-6">
            <FileSpreadsheet className="w-8 h-8 text-indigo-600" />
            <h1 className="text-3xl font-bold text-gray-800">OCI Metrics to Excel</h1>
          </div>

          <div className="mb-6">
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Compartment Name (Optional)
            </label>
            <input
              type="text"
              value={compartmentName}
              onChange={(e) => setCompartmentName(e.target.value)}
              placeholder="e.g., PRDAPPANLCOM"
              className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-colors"
            />
          </div>

          <div className="mb-6">
            <h3 className="text-lg font-semibold text-gray-700 mb-3">Upload CSV Files</h3>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <CsvUploadCard type="cpuMean" title="CPU Mean" color="blue" />
              <CsvUploadCard type="cpuMax" title="CPU Max" color="purple" />
              <CsvUploadCard type="memoryMean" title="Memory Mean" color="green" />
              <CsvUploadCard type="memoryMax" title="Memory Max" color="orange" />
            </div>
          </div>

          <button
            onClick={generateExcel}
            disabled={!Object.values(csvFiles).some(csv => csv !== null)}
            className="w-full bg-indigo-600 text-white py-4 px-6 rounded-lg font-medium hover:bg-indigo-700 disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center justify-center gap-2 transition-colors text-lg shadow-lg"
          >
            <Download className="w-6 h-6" />
            Generate Excel File
          </button>

          <div className="mt-6 bg-gray-50 rounded-lg p-4 max-h-64 overflow-y-auto">
            <h3 className="font-semibold text-gray-700 mb-2">Activity Log</h3>
            {logs.length === 0 ? (
              <p className="text-gray-500 text-sm">No activity yet. Upload CSV files to get started.</p>
            ) : (
              <div className="space-y-1">
                {logs.map((log, idx) => (
                  <div key={idx} className={`text-sm flex gap-2 ${
                    log.type === 'error' ? 'text-red-600' :
                    log.type === 'success' ? 'text-green-600' :
                    'text-gray-600'
                  }`}>
                    <span className="text-gray-400 text-xs">{log.time}</span>
                    <span>{log.message}</span>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="mt-6 p-4 bg-blue-50 rounded-lg">
            <h4 className="font-semibold text-blue-900 mb-2">How it works:</h4>
            <ol className="text-sm text-blue-800 space-y-1 list-decimal list-inside">
              <li>Enter compartment name (optional)</li>
              <li>Upload CSV files from OCI dashboard (at least one)</li>
              <li>Click "Generate Excel File" to create a new Excel file</li>
              <li>The Excel will contain all instances found in the CSVs with their latest metrics</li>
              <li>Repeat for other compartments and consolidate manually if needed</li>
            </ol>
            <div className="mt-3 p-3 bg-blue-100 rounded">
              <p className="text-xs text-blue-900 font-medium">
                <strong>Note:</strong> The app automatically extracts all instance names from your CSV files 
                and picks the latest metric value for each instance.
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default OCIMetricsUpdater;