// Test-Skript für das Excel-Layout
const ExcelJS = require('exceljs');

async function createTestExcel() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Test');
  
  // Header in Zeile 3
  ws.getCell('A3').value = 'Produkt-ID';
  ws.getCell('C3').value = 'Materialkurztext';
  ws.getCell('E3').value = 'Her.-Artikelnummer';
  ws.getCell('G3').value = 'Fert./Prüfhinweis';
  ws.getCell('I3').value = 'Werkstoff';
  ws.getCell('K3').value = 'Nettogewicht';
  ws.getCell('M3').value = 'Länge';
  ws.getCell('O3').value = 'Breite';
  ws.getCell('Q3').value = 'Höhe';
  
  // Daten ab Zeile 4
  ws.getCell('A4').value = 'A2V12345';
  ws.getCell('C4').value = 'Test Material';
  ws.getCell('E4').value = 'ART-001';
  ws.getCell('G4').value = 'Standard';
  ws.getCell('I4').value = 'Stahl';
  ws.getCell('K4').value = 1.5;
  ws.getCell('M4').value = 100;
  ws.getCell('O4').value = 50;
  ws.getCell('Q4').value = 25;
  
  ws.getCell('A5').value = 'A2V67890';
  ws.getCell('C5').value = 'Test Material 2';
  ws.getCell('E5').value = 'ART-002';
  ws.getCell('G5').value = 'Premium';
  ws.getCell('I5').value = 'Aluminium';
  ws.getCell('K5').value = 0.8;
  ws.getCell('M5').value = 80;
  ws.getCell('O5').value = 40;
  ws.getCell('Q5').value = 20;
  
  // Speichern
  await wb.xlsx.writeFile('test-input.xlsx');
  console.log('Test-Excel-Datei erstellt: test-input.xlsx');
}

createTestExcel().catch(console.error); 