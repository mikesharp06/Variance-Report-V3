document.getElementById('fileUpload').addEventListener('change', function(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = async function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(data.buffer);
            const worksheet = workbook.getWorksheet(1);

            // Delete rows 4, 5, and 6
            worksheet.spliceRows(4, 3);

            // Define the unmergeCells function
            const unmergeCells = (range) => {
                if (worksheet.getCell(range.split(':')[0]).isMerged) {
                    worksheet.unMergeCells(range);
                }
            };

            // Unmerge cells in Rows 1, 2, and 3 if they are merged from A to D
            unmergeCells('A1:D1');
            unmergeCells('A2:D2');
            unmergeCells('A3:D3');

            // Merge cells in Rows 1, 2, and 3 from A to I
            worksheet.mergeCells('A1:I1');
            worksheet.mergeCells('A2:I2');
            worksheet.mergeCells('A3:I3');

            // Center align the merged cells in Rows 1, 2, and 3
            worksheet.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };

            // Wrap text for all cells in Column J
            worksheet.getColumn('J').eachCell((cell) => {
                cell.alignment = { wrapText: true };
                cell.border = {};
            });

            // Set the header for Column J
            const varianceHeader = worksheet.getCell('J6');
            varianceHeader.value = 'Variance Comments';

            // Apply styles to the header
            varianceHeader.font = { bold: true };
            varianceHeader.alignment = { horizontal: 'center', vertical: 'middle' };
            varianceHeader.border = { bottom: { style: 'thin' } };

            // Set the width of Column J
            worksheet.getColumn('J').width = 40;

            worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                if (rowNumber >= 9) {
                    row.eachCell((cell) => {
                        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                    });
                }
            });

            // Remove cell with the text "Created on:"
            worksheet.eachRow((row) => {
                row.eachCell((cell) => {
                    if (cell.value === 'Created on:') {
                        cell.value = null;
                    }
                });
            });

            // Get the value of cell A1 to use as the filename
            const fileName = worksheet.getCell('A1').value ? `${worksheet.getCell('A1').value} - Modified` : 'modified';

            // Output the modified content to the console
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const downloadLink = document.getElementById('downloadLink');
            const downloadButton = document.getElementById('downloadButton');
            downloadButton.href = url;
            downloadButton.download = `${fileName}.xlsx`;
            downloadLink.style.display = 'block';
        };
        reader.readAsArrayBuffer(file);
    }
});