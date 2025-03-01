document.getElementById('uploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();
  
    const fileInput = document.getElementById('excelFile');
    const errorDiv = document.getElementById('error');
    const actTableBody = document.getElementById('actSummaryTable').querySelector('tbody');
    const secretaryTableBody = document.getElementById('secretarySummaryTable').querySelector('tbody');
  
    // Limpiar resultados anteriores y errores
    errorDiv.textContent = '';
    actTableBody.innerHTML = '';
    secretaryTableBody.innerHTML = '';
  
    if (!fileInput.files || fileInput.files.length === 0) {
      errorDiv.textContent = 'Por favor, selecciona un archivo Excel';
      return;
    }
  
    const file = fileInput.files[0];
  
    try {
      // Leer el archivo como array buffer
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet);
  
      // Verificar que las columnas esperadas existan
      const requiredColumns = ['Nro de Expediente', 'Acto Procesal', 'Secretario'];
      const firstRow = data[0];
      const missingColumns = requiredColumns.filter(col => !(col in firstRow));
      if (missingColumns.length > 0) {
        throw new Error(`Faltan las columnas: ${missingColumns.join(', ')}`);
      }
  
      // Calcular conteos
      const actSummary = {};
      const secretarySummary = {};
  
      data.forEach(row => {
        const act = row['Acto Procesal'] || 'Sin clasificar';
        const secretary = row['Secretario'] || 'Sin secretario';
        actSummary[act] = (actSummary[act] || 0) + 1;
        secretarySummary[secretary] = (secretarySummary[secretary] || 0) + 1;
      });
  
      // Mostrar resultados en las tablas
      Object.entries(actSummary).forEach(([act, count]) => {
        const row = document.createElement('tr');
        row.innerHTML = `<td>${act}</td><td>${count}</td>`;
        actTableBody.appendChild(row);
      });
  
      Object.entries(secretarySummary).forEach(([secretary, count]) => {
        const row = document.createElement('tr');
        row.innerHTML = `<td>${secretary}</td><td>${count}</td>`;
        secretaryTableBody.appendChild(row);
      });
  
      document.getElementById('results').style.display = 'block';
    } catch (error) {
      errorDiv.textContent = `Error al procesar el archivo: ${error.message}`;
      console.error('Error:', error);
    }
  });