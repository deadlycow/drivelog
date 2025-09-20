import React, { useState } from 'react';
import '../../css/excelhandler.css'
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';


const ExcelHandler = () => {
  const [data, setData] = useState([]);
  const [tempData, setTempData] = useState([]);
  const [fileName, setFileName] = useState('');

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        setData(jsonData);
        // console.log('Data loaded:', jsonData);
      } catch (error) {
        console.error('Error reading file:', error);
        alert('Fel vid läsning av fil');
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const sortList = () => {
    const sortList = data.flatMap((row, rowIndex) =>
      Object.entries(row)
        .filter(([key, value]) => typeof value === 'string' && value.toLowerCase().includes('led'))
        .map(([key, value]) => ({
          rowIndex,
          key,
          original: value
        }))
    )
    setTempData(sortList)
  }

  const manipulateDataLocally = () => {
    if (data.length === 0) return;

    const manipulatedData = data.map((row, index) => ({
      ...row,
      // Exempel på manipulering
      RowNumber: index + 1,
      ProcessedAt: new Date().toISOString(),
      // Lägg till dina egna manipuleringar här
    }));

    setData(manipulatedData);
  };

  // Skapa och ladda ner ny Excel-fil
  const downloadExcel = () => {
    if (data.length === 0) {
      alert('Ingen data att exportera');
      return;
    }

    try {
      const worksheet = XLSX.utils.json_to_sheet(data);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Processed Data');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      const newFileName = fileName ?
        `processed_${fileName}` :
        `processed_data_${new Date().toISOString().split('T')[0]}.xlsx`;

      saveAs(blob, newFileName);
    } catch (error) {
      console.error('Error creating Excel file:', error);
      alert('Fel vid skapande av Excel-fil');
    }
  };

  return (
    <div className='container'>
      <h2>Excel Hantering</h2>

      {/* Fil uppladdning */}
      <div style={{ marginBottom: '20px' }}>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          style={{ marginBottom: '10px' }}
        />
        {fileName && <p>Fil laddad: {fileName}</p>}
      </div>

      {data.length > 0 && (
        <>
          <div className='table-container'>
            <table>
              <thead>
                <tr>
                  {/* Ta första objektet i data och använd dess keys som kolumnnamn */}
                  {data.length > 0 &&
                    Object.keys(data[0]).map((key, i) => (
                      <th key={i}>
                        {key}
                      </th>
                    ))}
                </tr>
              </thead>
              <tbody>
                {(tempData.length > 0 ? tempData : data).map((row, i) => (
                  <tr key={i}>
                    {Object.values(row).map((value, j) => (
                      <td key={j} >
                        {value}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
        <button
          className='btn-man'
          onClick={manipulateDataLocally}
          disabled={data.length === 0} >
          Manipulera Data Lokalt
        </button>
        <button
          className='btn-down'
          onClick={downloadExcel}
          disabled={data.length === 0} >
          Ladda ner Excel (Lokalt)
        </button>
        <button
          className='btn-entries'
          onClick={sortList}
          disabled={data.length === 0} >
          Filter
        </button>
      </div>
    </div>
  );
};

export default ExcelHandler;