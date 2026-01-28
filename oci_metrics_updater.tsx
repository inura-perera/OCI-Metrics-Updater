import React, { useState } from 'react';
import { Upload, Download, RefreshCw, FileSpreadsheet, X } from 'lucide-react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

const OCIMetricsUpdater = () => {
  const [excelData, setExcelData] = useState(null);
  const [csvFiles, setCsvFiles] = useState({
    cpuMean: null,
    cpuMax: null,
    memoryMean: null,
    memoryMax: null
  });
  const [logs, setLogs] = useState([]);
  const [processing, setProcessing] = useState(false);

  const addLog = (message, type = 'info') => {
    setLogs(prev => [...prev, { message, type, time: new Date().toLocaleTimeString() }]);
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        setExcelData({ workbook, worksheet, sheetName });
        addLog('Excel file loaded successfully', 'success');
      } catch (error) {
        addLog('Error loading Excel file: ' + error.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
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
        addLog(`${type.replace(/([A-Z])/g, ' $1').toUpperCase()} CSV loaded: ${results.data.length} rows`, 'success');
      },
      error: (error) => {
        addLog(`Error parsing ${type} CSV: ` + error.message, 'error');
      }
    });
  };

  const clearCsv = (type) => {
    setCsvFiles(prev => ({ ...prev, [type]: null }));
    addLog(`${type.replace(/([A-Z])/g, ' $1').toUpperCase()} CSV removed`, 'info');
  };

  const getMetricValue = (csvData, instanceName) => {
    if (!csvData || csvData.length === 0) return null;

    // Find the row that matches the instance name
    for (const row of csvData) {
      // Check each column to find the instance name
      for (const [key, value] of Object.entries(row)) {
        if (value && value.toString().trim() === instanceName) {
          // Found the matching row, now get the latest metric value
          // Get all numeric columns (these are the metric values with timestamps)
          const metricColumns = Object.entries(row)
            .filter(([k, v]) => k !== 'group' && typeof v === 'number')
            .sort((a, b) => b[0].localeCompare(a[0])); // Sort by timestamp desc

          if (metricColumns.length > 0) {
            // Return the most recent value (first in sorted array)
            return metricColumns[0][1];
          }
        }
      }
    }
    return null;
  };

  const updateMetrics = () => {
    if (!excelData) {
      addLog('Please upload Excel file', 'error');
      return;
    }

    const hasAnyCsv = Object.values(csvFiles).some(csv => csv !== null);
    if (!hasAnyCsv) {
      addLog('Please upload at least one CSV file', 'error');
      return;
    }

    setProcessing(true);
    addLog('Starting metrics update...', 'info');

    try {
      const { worksheet } = excelData;
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      
      // Column indices: B=1 (Instance Name), G=6 (CPU Mean), H=7 (CPU Max), I=8 (RAM Mean), J=9 (RAM Max)
      const instanceCol = 1;
      let updatedCount = 0;

      // Iterate through Excel rows starting from row 4 (index 3)
      for (let row = 3; row <= range.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: instanceCol });
        const cell = worksheet[cellAddress];
        
        if (!cell || !cell.v) continue;
        
        const instanceName = cell.v.toString().trim();
        let instanceUpdated = false;

        // Update CPU Mean (Column G)
        if (csvFiles.cpuMean) {
          const value = getMetricValue(csvFiles.cpuMean, instanceName);
          if (value !== null) {
            const cellAddr = XLSX.utils.encode_cell({ r: row, c: 6 });
            worksheet[cellAddr] = { t: 'n', v: value / 100, z: '0%' };
            instanceUpdated = true;
          }
        }

        // Update CPU Max (Column H)
        if (csvFiles.cpuMax) {
          const value = getMetricValue(csvFiles.cpuMax, instanceName);
          if (value !== null) {
            const cellAddr = XLSX.utils.encode_cell({ r: row, c: 7 });
            worksheet[cellAddr] = { t: 'n', v: value / 100, z: '0%' };
            instanceUpdated = true;
          }
        }

        // Update Memory Mean (Column I)
        if (csvFiles.memoryMean) {
          const value = getMetricValue(csvFiles.memoryMean, instanceName);
          if (value !== null) {
            const cellAddr = XLSX.utils.encode_cell({ r: row, c: 8 });
            worksheet[cellAddr] = { t: 'n', v: value / 100, z: '0%' };
            instanceUpdated = true;
          }
        }

        // Update Memory Max (Column J)
        if (csvFiles.memoryMax) {
          const value = getMetricValue(csvFiles.memoryMax, instanceName);
          if (value !== null) {
            const cellAddr = XLSX.utils.encode_cell({ r: row, c: 9 });
            worksheet[cellAddr] = { t: 'n', v: value / 100, z: '0%' };
            instanceUpdated = true;
          }
        }

        if (instanceUpdated) {
          updatedCount++;
          addLog(`✓ Updated: ${instanceName}`, 'success');
        }
      }

      addLog(`Update complete! ${updatedCount} instances updated`, 'success');
    } catch (error) {
      addLog('Error updating metrics: ' + error.message, 'error');
    } finally {
      setProcessing(false);
    }
  };

  const downloadExcel = () => {
    if (!excelData) {
      addLog('No Excel data to download', 'error');
      return;
    }

    try {
      const { workbook } = excelData;
      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `OCI_Metrics_Updated_${new Date().toISOString().split('T')[0]}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      addLog('Excel file downloaded successfully', 'success');
    } catch (error) {
      addLog('Error downloading Excel: ' + error.message, 'error');
    }
  };

  const CsvUploadCard = ({ type, title, color }) => (
    <div className={`border-2 border-dashed rounded-lg p-4 hover:border-${color}-400 transition-colors ${csvFiles[type] ? 'border-' + color + '-400 bg-' + color + '-50' : 'border-gray-300'}`}>
      <label className="flex flex-col items-center cursor-pointer">
        <Upload className={`w-8 h-8 mb-2 ${csvFiles[type] ? 'text-' + color + '-600' : 'text-gray-400'}`} />
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
          <span className={`text-xs text-${color}-600 font-medium`}>
            ✓ {csvFiles[type].length} rows
          </span>
          <button
            onClick={() => clearCsv(type)}
            className="text-gray-500 hover:text-red-600"
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
            <h1 className="text-3xl font-bold text-gray-800">OCI Metrics Updater</h1>
          </div>

          <div className="mb-6">
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 hover:border-indigo-400 transition-colors">
              <label className="flex flex-col items-center cursor-pointer">
                <Upload className="w-12 h-12 text-gray-400 mb-2" />
                <span className="text-sm font-medium text-gray-700 mb-2">Upload Excel Template</span>
                <span className="text-xs text-gray-500 text-center">Your daily Excel sheet (.xlsx)</span>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleExcelUpload}
                  className="hidden"
                />
              </label>
              {excelData && (
                <div className="mt-3 text-center text-sm text-green-600 font-medium">
                  ✓ Excel loaded
                </div>
              )}
            </div>
          </div>

          <div className="mb-6">
            <h3 className="text-lg font-semibold text-gray-700 mb-3">Upload CSV Files for Compartment</h3>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <CsvUploadCard type="cpuMean" title="CPU Mean" color="blue" />
              <CsvUploadCard type="cpuMax" title="CPU Max" color="purple" />
              <CsvUploadCard type="memoryMean" title="Memory Mean" color="green" />
              <CsvUploadCard type="memoryMax" title="Memory Max" color="orange" />
            </div>
          </div>

          <div className="flex gap-4 mb-6">
            <button
              onClick={updateMetrics}
              disabled={!excelData || processing}
              className="flex-1 bg-indigo-600 text-white py-3 px-6 rounded-lg font-medium hover:bg-indigo-700 disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center justify-center gap-2 transition-colors"
            >
              <RefreshCw className={`w-5 h-5 ${processing ? 'animate-spin' : ''}`} />
              {processing ? 'Processing...' : 'Update Metrics'}
            </button>

            <button
              onClick={downloadExcel}
              disabled={!excelData}
              className="flex-1 bg-green-600 text-white py-3 px-6 rounded-lg font-medium hover:bg-green-700 disabled:bg-gray-300 disabled:cursor-not-allowed flex items-center justify-center gap-2 transition-colors"
            >
              <Download className="w-5 h-5" />
              Download Updated Excel
            </button>
          </div>

          <div className="bg-gray-50 rounded-lg p-4 max-h-64 overflow-y-auto">
            <h3 className="font-semibold text-gray-700 mb-2">Activity Log</h3>
            {logs.length === 0 ? (
              <p className="text-gray-500 text-sm">No activity yet. Upload files to get started.</p>
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
            <h4 className="font-semibold text-blue-900 mb-2">How to use:</h4>
            <ol className="text-sm text-blue-800 space-y-1 list-decimal list-inside">
              <li>Upload your Excel template first</li>
              <li>Upload the 4 CSV files from OCI dashboard (one compartment at a time)</li>
              <li>Click "Update Metrics" to populate the cells</li>
              <li>Repeat steps 2-3 for other compartments if needed</li>
              <li>Download the final updated Excel file</li>
            </ol>
            <p className="text-xs text-blue-700 mt-2">
              <strong>Tip:</strong> You can upload any combination of CSV files. If you only have CPU data, upload just those files.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default OCIMetricsUpdater;