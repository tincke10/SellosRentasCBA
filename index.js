const formConversion = document.getElementById('formConversion');

formConversion.addEventListener('submit', (event) => {

    event.preventDefault();

    const archivoXLSX = document.getElementById('archivoExcel').files[0];

    if (!archivoXLSX) {
        Swal.fire({
            title: "Error!",
            text: "Debe subir un archivo Excel para procesar",
            icon: "error"
        });
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {

        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        //hoja 1
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Obtén el rango de celdas utilizadas en la hoja (ajustando según las columnas y filas deseadas)
        const startRow = 2; // Fila 3 
        const endRow = XLSX.utils.decode_range(worksheet['!ref']).e.r; //hasta que no haya mas filas con datos
        const startCol = XLSX.utils.decode_col("B"); // Columna B
        const endCol = XLSX.utils.decode_col("M"); // Columna M

        let txtData = '';
        let count = 1;

        // Itera sobre cada fila y obtengo las columnas específicas que necesito
        for (let rowNum = startRow; rowNum <= endRow; rowNum++) {

            //verifico que no sea una fila vacia
            if (!worksheet[XLSX.utils.encode_cell({ c: XLSX.utils.decode_col("B"), r: rowNum })] ||
                !worksheet[XLSX.utils.encode_cell({ c: XLSX.utils.decode_col("C"), r: rowNum })] ||
                !worksheet[XLSX.utils.encode_cell({ c: XLSX.utils.decode_col("D"), r: rowNum })]) {
                break
            }

            let rowArray = [];
            rowArray.push(count++);

            for (let colNum = startCol; colNum <= endCol; colNum++) {

                //cada celda
                const cellAddress = { c: colNum, r: rowNum };
                const cellRef = XLSX.utils.encode_cell(cellAddress);
                const cell = worksheet[cellRef];

                // Añade el valor de la celda al array de la fila
                rowArray.push(cell ? cell.v : '');
            }

            //modifico el rowArray y lo formateo a como debería ser el .txt final
            const arrayFormateado = arrayFormatter(rowArray);

            // Convierte el array a una cadena de texto separada por tabulaciones
            txtData += arrayFormateado + '\n';
        }

        // Crear un Blob con los datos de texto
        txtData = txtData.slice(0, -1);
        const blob = new Blob([txtData], { type: 'text/plain' });

        // Crear un enlace para descargar el archivo
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = archivoXLSX.name.slice(0, -4) + '.txt';
        link.click();
    };


    //llamo al reader con el archivo leído
    reader.readAsArrayBuffer(archivoXLSX);

});

//funciona que formatea y devuelve cada fila del archivo
function arrayFormatter(array) {

    // Verifica si el valor tiene un solo dígito y agrega un 0
    const numFila = array[0] < 10 ? '0' + array[0].toString() : array[0].toString();

    const dia = array[1] < 10 ? '0' + array[1].toString() : array[1].toString();
    const mes = array[2] < 10 ? '0' + array[2].toString() : array[2].toString();

    const fecha = dia + '/' + mes + '/' + array[3];

    const cuil = array[12];

    //queda con 2 dígitos decimales
    let monto = array[6].toFixed(2).toString();
    monto = monto.replace('.', ',');
    const valor = monto.split(',')[0];
    const centavos = monto.split(',')[1];

    //queda con 2 dígitos decimales
    let alicuota = (array[7] * 100).toFixed(2).toString();
    alicuota = alicuota.replace('.', ',');

    let impuesto = array[8].toFixed(2).toString();
    impuesto = impuesto.replace('.', ',');
    const valorImpuesto = impuesto.split(',')[0];
    const centavosImpuesto = impuesto.split(',')[1];

    return numFila + '000001001' + fecha + fecha + '5' + cuil + valorPart(valor) + ',' + centavos + '0' + alicuota + '00000000000,0000000000000000000000000000000000000,' + valorPart(valorImpuesto) + ',' + centavosImpuesto;

}

//funciona que crea un string de 10 digitos, pero que se "acomoda" a la cantidad de digitos del valor
function valorPart(valor) {

    const zeros = '0000000000';
    const valorPart = zeros.substring(0, 10 - valor.length) + valor;
    return valorPart;

}