function HV() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(hoja.getName() !== 'HV'){
    return;
  }
  const carga = hoja.getRange('N5');
  const diagonal = hoja.getRange('N6');
  const dureza = hoja.getRange('N7');

  // Verifica cuál de las celdas está vacía y realiza el cálculo correspondiente
  if (dureza.getValue() === "") {
    // Calcula dureza si N7 está vacía
    dureza.setValue(
      1.8453 * (carga.getValue() / (diagonal.getValue() ** 2))
    );
  } else if (carga.getValue() === "") {
    // Calcula carga si N5 está vacía
    carga.setValue(
      (dureza.getValue() * (diagonal.getValue() ** 2)) / 1.8453
    );
  } else if (diagonal.getValue() === "") {
    // Calcula diagonal si N6 está vacía
    diagonal.setValue(
      Math.sqrt(carga.getValue() / dureza.getValue())
    );
  }
}