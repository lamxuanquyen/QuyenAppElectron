const { ipcRenderer } = require('electron');
const ExcelJS = require('exceljs');

let fileLists = []; 
async function loadFileLists(file) {
    return new Promise((resolve, reject) => {
        // const file = document.getElementById('myFile').files[0];
        const reader = new FileReader();

        reader.onload = async function(e) {
            try {
                const buffer = new Uint8Array(e.target.result);
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(buffer);

                const worksheet = workbook.getWorksheet(1);
                worksheet.eachRow((row, rowNumber) => {
                    fileLists.push(row.getCell(1).value.toString().trim());
                });

                resolve(fileLists);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = function(e) {
            reject(e);
        };

        if (file) {
            reader.readAsArrayBuffer(file);
        } else {
            reject(new Error('No file selected.'));
        }
    });
}

document.getElementById('sourceBtn').addEventListener('click', () => {
    ipcRenderer.invoke('open-dialog', 'openDirectory')
    .then(data => {
        document.getElementById('sourceInput').value = data;
    });
});

document.getElementById('destBtn').addEventListener('click', () => {
    ipcRenderer.invoke('open-dialog', 'openDirectory')
    .then(data => {
        document.getElementById('destInput').value = data;
    });
});

document.getElementById('copyBtn').addEventListener('click', async () => {
    let sourcePath = document.getElementById('sourceInput').value.trim();
    let destPath = document.getElementById('destInput').value.trim();
    let fileType = document.getElementById('fileType').value;
    let file = document.getElementById('myFile').files[0];

    if (!sourcePath || !destPath) {
        alert('Please select source and destination folders.');
        return;
    }

    if (!fileType) {
        fileType = '';
    }

    if (!file) {
        alert('Please select a file.');
        return;
    }

    await loadFileLists(file);

    ipcRenderer.invoke('copy-file', sourcePath, destPath, fileLists, fileType)
    .then(results => {
        // Tính toán số lượng kết quả
        let successCount = 0;
        let notFoundCount = 0;
        let alreadyExistsCount = 0;

        results.forEach(result => {
            switch (result.status) {
                case 'File copied successfully':
                successCount++;
                break;
                case 'File not found':
                notFoundCount++;
                break;
                case 'File already exists in target folder':
                alreadyExistsCount++;
                break;
            }
        });

        // Cập nhật nội dung của thẻ <p id="result">
        document.getElementById('result').textContent = 
        `${successCount} file copy thành công, ` +
        `${notFoundCount} file không tìm thấy, ` +
        `${alreadyExistsCount} file đã có sẵn ở target folder. ` +
        `Xem chi tiết ở file outputCopy.xlsx`;
    });

    fileLists = [];
});

document.getElementById('moveBtn').addEventListener('click', async () => {
    let sourcePath = document.getElementById('sourceInput').value.trim();
    let destPath = document.getElementById('destInput').value.trim();
    let fileType = document.getElementById('fileType').value;
    let file = document.getElementById('myFile').files[0];

    if (!sourcePath || !destPath) {
        alert('Please select source and destination folders.');
        return;
    }

    if (!fileType) {
        fileType = '';
    }

    if (!file) {
        alert('Please select a file.');
        return;
    }
    
    await loadFileLists(file);

    ipcRenderer.invoke('move-file', sourcePath, destPath, fileLists, fileType)
    .then(results => {
        // Tính toán số lượng kết quả
        let successCount = 0;
        let notFoundCount = 0;
        let errorCount = 0;

        results.forEach(result => {
            switch (result.status) {
                case 'File moved successfully':
                successCount++;
                break;
                case 'File not found':
                notFoundCount++;
                break;
                case 'File moved error':
                errorCount++;
                break;
            }
        });

        // Cập nhật nội dung của thẻ <p id="result">
        document.getElementById('result').textContent = 
        `${successCount} file move thành công, ` +
        `${notFoundCount} file không tìm thấy, ` +
        `${errorCount} file đã bị lỗi khi move. ` +
        `Xem chi tiết ở file outputMove.xlsx`;
    });

    fileLists = [];
});

document.getElementById('select-folder').addEventListener('click', () => {
    ipcRenderer.invoke('open-dialog', 'openDirectory')
    .then(data => {
        document.getElementById('pdf-folder-path').value = data;
    });
});

document.getElementById('merge-pdfs').addEventListener('click', () => {
    let folderPath = document.getElementById('pdf-folder-path').value.trim();
    
    if (!folderPath) {
        alert('Please select a folder.');
        return;
    }

    ipcRenderer.invoke('merge-pdfs', folderPath)
    .then(message => {
        document.getElementById('pdf-folder-path').value = '';
        document.getElementById('resultMerger').textContent = message;
    });
});

document.getElementById('pathBtn').addEventListener('click', () => {
    ipcRenderer.invoke('open-dialog', 'openDirectory')
    .then(data => {
        document.getElementById('folderPath').value = data;
    });
});

document.getElementById('getPrint').addEventListener('click', () => {
    ipcRenderer.invoke('get-printers').then(printers => {
      const printerSelect = document.getElementById('printerName');
      // Xóa các option cũ trước khi thêm mới
      printerSelect.innerHTML = '';
      printers.forEach(printer => {
        const option = document.createElement('option');
        option.value = printer.name;
        option.textContent = printer.displayName || printer.name;
        // Đánh dấu máy in mặc định
        if (printer.isDefault) {
            option.selected = true;
        }
        printerSelect.appendChild(option);
      });
    });
});

document.getElementById('printBtn').addEventListener('click', async () => {
    let folderPath = document.getElementById('folderPath').value.trim();
    let printerName = document.getElementById('printerName').value.trim();
    let paperSize = document.getElementById('paperSize').value;
    let file = document.getElementById('excelFile').files[0];

    if (!file) {
        alert('Please select a file.');
        return;
    }

    if (!folderPath) {
        alert('Please select a folder.');
        return;
    }

    if (!printerName) {
        alert('Please select a printer.');
        return;
    }

    if (!paperSize) {
        alert('Please select a paper size.');
        return;
    }

    await loadFileLists(file);

    ipcRenderer.invoke('print-pdf-files', folderPath, printerName,paperSize, fileLists)
    .then(results => {
        let successCount = 0;
        let notFoundCount = 0;
        let errorCount = 0;

        results.forEach(result => {
            switch (result.status) {
                case 'Printed successfully':
                successCount++;
                break;
                case 'File not found':
                notFoundCount++;
                break;
                case 'Failed to print':
                errorCount++;
                break;
            }
        });

        // Cập nhật nội dung của thẻ <p id="resultPrint">
        document.getElementById('resultPrint').textContent = 
        `${successCount} file in thành công, ` +
        `${notFoundCount} file không tìm thấy, ` +
        `${errorCount} file đã bị lỗi khi in. ` +
        `Xem chi tiết ở file outputPrint.xlsx`;
    })
    .catch(error => {
        // Trường hợp xảy ra lỗi
        document.getElementById('resultPrint').textContent = `Lỗi khi in: ${error.message}`;
    });

    fileLists = [];
});