const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const fs = require('fs-extra');
const ExcelJS = require('exceljs');
const { PDFDocument } = require('pdf-lib');

function createWindow () {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  win.setMenu(null);
  // Menu.setApplicationMenu(null);

  win.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  app.quit();
});

ipcMain.handle('open-dialog', async (event, dialogType) => {
  const result = await dialog.showOpenDialog({ properties: [dialogType] });
  return result.filePaths[0];
});

ipcMain.handle('copy-file', async (event, sourcePath, destPath, fileLists, fileType) => {
  const results = [];

  // console.log("cac tham so copy la : " + sourcePath, destPath, fileLists, fileType)

  if (!fs.existsSync(sourcePath)) {
    moveResults.push({file: sourcePath, status: 'Source folder not found'});
    return moveResults;
  }

  if (!fs.existsSync(destPath)) {
    moveResults.push({file: destPath, status: 'Destination folder not found'});
    return moveResults;
  }

  if (!fileType) {
    fileType = '';
  }

  if (!fileLists) {
    fileLists = [];
  }

  if (fileLists.length === 0) {
    moveResults.push({file: 'No files selected', status: 'No files selected'});
    return moveResults;
  }
  
  const files = await fs.readdir(sourcePath);

  for (const fileName of fileLists) {
    const filesToCopy = files.filter(file => path.parse(file).name === fileName && (fileType === '' || path.extname(file) === fileType));

    if (filesToCopy.length === 0) {
      results.push({file: `${fileName}${fileType}`, status: 'File not found'});
      continue;
    } 

    await Promise.all(filesToCopy.map(async file => {
      // Kiểm tra xem tệp không tồn tại trong thư mục đích hay chưa   
      if (!fs.existsSync(path.join(destPath, file))) {
        await fs.copy(path.join(sourcePath, file), path.join(destPath, file));
        results.push({file: file, status: 'File copied successfully'});
      }else {
        results.push({file: file, status: 'File already exists in target folder'});
      }
    }));
    
  }

  // Tạo một workbook mới
  const workbook = new ExcelJS.Workbook();

  // Thêm một worksheet mới vào workbook
  const worksheet = workbook.addWorksheet('Results');

  // Định nghĩa các cột cho worksheet
  worksheet.columns = [
    {header: 'File Name', key: 'file', width: 50},
    {header: 'Result', key: 'status', width: 50}
  ];

  // Thêm các hàng vào worksheet
  results.forEach(result => {
    worksheet.addRow(result);
  });

  // Ghi workbook vào file
  await workbook.xlsx.writeFile('outputCopy.xlsx');

  // console.log('Results written to output.xlsx');

  return results;
});

ipcMain.handle('move-file', async (event, sourcePath, destPath, fileLists, fileType) => {
  const moveResults = [];

  // console.log("cac tham so move la : " + sourcePath, destPath, fileLists, fileType)

  if (!fs.existsSync(sourcePath)) {
    moveResults.push({file: sourcePath, status: 'Source folder not found'});
    return moveResults;
  }

  if (!fs.existsSync(destPath)) {
    moveResults.push({file: destPath, status: 'Destination folder not found'});
    return moveResults;
  }

  if (!fileType) {
    fileType = '';
  }

  if (!fileLists) {
    fileLists = [];
  }

  if (fileLists.length === 0) {
    moveResults.push({file: 'No files selected', status: 'No files selected'});
    return moveResults;
  }

  const files = await fs.readdir(sourcePath);

  for (const fileName of fileLists) {
    const filesToMove = files.filter(file => path.parse(file).name === fileName && (fileType === '' || path.extname(file) === fileType));

    if (filesToMove.length === 0) {
      moveResults.push({file: `${fileName}${fileType}`, status: 'File not found'});
      continue;
    } 

    await Promise.all(filesToMove.map(async file => {
      // Kiểm tra xem tệp không tồn tại trong thư mục đích hay chưa   
      if (!fs.existsSync(path.join(destPath, file))) {
        await fs.move(path.join(sourcePath, file), path.join(destPath, file));
        moveResults.push({file: file, status: 'File moved successfully'});
      }else {
        // fs.unlinkSync(path.join(sourcePath, file));
        let filePath = path.join(sourcePath, file);
        fs.unlink(filePath, (err) => {
            if (err) {
                if (err.code === 'ENOENT') {
                    // console.log(`File ${filePath} khong ton tai.`);
                    moveResults.push({file: filePath, status: 'File not found'});
                } else {
                    console.log(`Loi khi xoa file ${filePath}:`, err);
                    moveResults.push({file: filePath, status: 'File moved error'});
                }
            } else {
                // console.log(`da xoa file ${filePath}`);
                moveResults.push({file: filePath, status: 'File moved successfully'});
            }
        });
      }
    }));

  }

  // Tạo một workbook mới
  const workbook = new ExcelJS.Workbook();

  // Thêm một worksheet mới vào workbook
  const worksheet = workbook.addWorksheet('Results');

  // Định nghĩa các cột cho worksheet
  worksheet.columns = [
    {header: 'File Name', key: 'file', width: 50},
    {header: 'Result', key: 'status', width: 50}
  ];

  // Thêm các hàng vào worksheet
  moveResults.forEach(result => {
    worksheet.addRow(result);
  });

  // Ghi workbook vào file
  await workbook.xlsx.writeFile('outputMove.xlsx');

  return moveResults;
});

ipcMain.handle('merge-pdfs', async (event, folderPath) => {
  try {
    const mergedPdf = await PDFDocument.create();

    if (!fs.existsSync(folderPath)) {
      return 'Folder not found.';
    }

    if (fs.readdirSync(folderPath).length === 0) {
      return 'Folder is empty.';
    }

    if (!fs.statSync(folderPath).isDirectory()) {
      return 'Path is not a folder.';
    }
    
    const files = fs.readdirSync(folderPath).filter(file => file.endsWith('.pdf'));

    if (files.length === 0) {
      return 'No PDF files found in the folder.';
    }
    
    for (const file of files) {
      const filePath = path.join(folderPath, file);
      const pdfBytes = fs.readFileSync(filePath);
      const pdfDoc = await PDFDocument.load(pdfBytes);
      const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
      copiedPages.forEach((page) => {
        mergedPdf.addPage(page);
      });
    }
    
    const mergedPdfBytes = await mergedPdf.save();
    const outputPath = path.join(folderPath, 'merged.pdf');
    fs.writeFileSync(outputPath, mergedPdfBytes);
    
    return `PDFs merged successfully! . See output in ${outputPath}`;
  } catch (error) {
    console.error('Error merging PDFs:', error);
    return 'Failed to merge PDFs: ' + error.message;
  }
});

ipcMain.handle('get-printers', async () => {
  const win = new BrowserWindow({ show: false });
  try {
    const printers = await win.webContents.getPrintersAsync();
    return printers;
  } catch (error) {
    // Xử lý lỗi ở đây
    console.error('An error occurred while getting printers:', error);
  } finally {
    // Đóng cửa sổ BrowserWindow bất kể có lỗi hay không
    win.close();
  }
});


ipcMain.handle('print-pdf-files', async (event, folderPath, printerName,paperSize, fileLists) => {
  const printResults = [];
  let fileType = '.pdf';

  console.log("cac tham so print la : " + folderPath, printerName,paperSize, fileLists)

  if (!fs.existsSync(folderPath)) {
    printResults.push({ fileName: '', printResult: 'Folder not found' });
    return printResults;
  }

  if (fileLists.length === 0) {
    printResults.push({ fileName: '', printResult: 'No files to print' });
    return printResults;
  }

  if (!printerName) {
    printResults.push({ fileName: '', printResult: 'No printer selected' });
    return printResults;
  }

  if (!paperSize) {
    printResults.push({ fileName: '', printResult: 'No paper size selected' });
    return printResults;
  }


  for (const fileName of fileLists) {
    const filePath = path.join(folderPath, fileName+fileType);

    // Kiem tra xem file co ton tai khong
    if (!fs.existsSync(filePath)) {
      printResults.push({ fileName, printResult: 'File not found' });
      continue;
    }

    // // Kiem tra xem file co phai file PDF khong
    // if (!fileName.endsWith('.pdf')) {
    //   printResults.push({ fileName, printResult: 'Not a PDF file' });
    //   continue;
    // }

    // Tạo một window để in
    const win = new BrowserWindow({ show: false });
    await win.loadFile(filePath);

    // Thực hiện việc in
    win.webContents.print({
      silent: true,
      printBackground: true,
      deviceName: printerName,
      pageSize: paperSize
    }, (success, errorType) => {
      if (success) {
        printResults.push({ fileName, printResult: 'Printed successfully' });
      } else {
        printResults.push({ fileName, printResult: `Failed to print` });
      }
    });

    // Đợi cho đến khi việc in hoàn tất hoặc thất bại
    await new Promise(resolve => win.webContents.on('did-finish-load', resolve));

    win.close();
  }

  // Tạo một workbook mới
  const workbook = new ExcelJS.Workbook();

  // Thêm một worksheet mới vào workbook
  const worksheet = workbook.addWorksheet('Results');

  // Định nghĩa các cột cho worksheet
  worksheet.columns = [
    {header: 'File Name', key: 'fileName', width: 50},
    {header: 'Result', key: 'printResult', width: 50}
  ];

  // Thêm các hàng vào worksheet
  printResults.forEach(result => {
    worksheet.addRow(result);
  });

  // Ghi workbook vào file
  await workbook.xlsx.writeFile('outputPrint.xlsx');

  return printResults;
});