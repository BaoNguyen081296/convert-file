/* eslint-disable no-loop-func */
import fs from 'file-saver';
import { formatDate, TYPE, EXCEL_PASSWORD } from 'utils/utils';
import _isEmpty from 'lodash/isEmpty';
const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();
workbook.creator = 'DeHR';
workbook.created = new Date();
workbook.calcProperties.fullCalcOnLoad = true;
const worksheet = workbook.addWorksheet('JsonToExcel');
const styles = {
  protection: {
    locked: {
      locked: true,
      hidden: true
    },
    unlocked: {
      locked: false,
      hidden: false
    }
  },
  fillCellForm: {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {
      argb: 'b4d9c2'
    }
  },
  fillEmptyCellForm: {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {
      argb: 'f45c5c'
    }
  }
};

const worksheetAddRow = (ws, data, index) => ws.addRow(data);

export const exportFile = async ({ type = TYPE.TO_JSON, file }) => {
  try {
    if (type === TYPE.TO_JSON) handleExcelToJson(file);
    else {
      let secondFile = null;
      for (let i = file.length - 1; i >= 0; i--) {
        const reader = new FileReader();
        reader.readAsText(file[i].originFileObj);
        reader.addEventListener('load', async e => {
          if (i === 0) handleJsonToExcel(JSON.parse(reader.result), secondFile);
          else secondFile = await JSON.parse(reader.result);
        });
      }
    }
  } catch (error) {
    console.log('error: ', error);
  }
};

const handleJsonToExcel = (file, secondFile) => {
  transformDataToXLSX(file, worksheet, secondFile);
  workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    fs.saveAs(blob, 'JsonToExcel-' + formatDate(new Date(), 'YYYYMMDDHHmm') + '.xlsx');
  });
};

const handleExcelToJson = async file => {
  let rows = (await workbook.xlsx.load(file)).getWorksheet()._rows;
  const data = transformDataToJson(rows);
  const blob = new Blob([JSON.stringify(data)], { type: 'text/plain;charset=utf-8' });
  return fs.saveAs(blob, 'convertedFile.json');
};

const transformDataToXLSX = async (file, ws, secondFile) => {
  try {
    // col width
    const columnsWidth = [15, 15, 15, 50, 50];
    columnsWidth.forEach((item, index) => {
      if (item) worksheet.getColumn(index + 1).width = item;
    });

    let titleRow = [
      'First Key',
      'Second Key',
      'Third Key',
      'English',
      secondFile ? 'Vietnamese' : ''
    ];
    unProtectValueCell(worksheetAddRow(ws, titleRow));
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'A1' }];
    if (typeof file === 'object') {
      const data = Object.keys(file);
      data.forEach(d => {
        if (typeof file[d] === 'object') {
          // first Key
          const keys = Object.keys(file[d]);
          keys.forEach(item => {
            if (typeof file[d][item] === 'object') {
              // second key
              Object.keys(file[d][item]).forEach(i => {
                let data = [d, item, i, file[d][item][i], secondFile?.[d]?.[item][i] || ''];
                let row = worksheetAddRow(ws, data);
                unProtectValueCell(row);
              });
            } else {
              let data = [d, item, '', file[d][item], secondFile?.[d]?.[item] || ''];
              unProtectValueCell(worksheetAddRow(ws, data));
            }
          });
        } else {
          let data = [d, '', file[d], secondFile[d] || ''];
          unProtectValueCell(worksheetAddRow(ws, data));
        }
      });
    }
    await ws.protect(EXCEL_PASSWORD);
  } catch (error) {
    console.log('error: ', error);
    throw error;
  }
};

const transformDataToJson = rows => {
  try {
    console.log('rows: ', rows);
    let obj = {};
    rows.forEach(i => {
      const r = i.values;
      if (!_isEmpty(r)) {
        switch (r.length) {
          case 4:
            obj = { ...obj, [r[1]]: { ...obj[r[1]], [r[2]]: r[3] } };
            break;
          case 5:
            if (!_isEmpty(obj[r[1]])) {
              if (!_isEmpty(obj[r[1]][r[2]])) {
                obj = {
                  ...obj,
                  [r[1]]: {
                    ...obj[r[1]],
                    [r[2]]: {
                      ...obj[r[1]][r[2]],
                      [r[3]]: r[4]
                    }
                  }
                };
              } else {
                obj = {
                  ...obj,
                  [r[1]]: {
                    ...obj[r[1]],
                    [r[2]]: {
                      [r[3]]: r[4]
                    }
                  }
                };
              }
            } else {
              obj = {
                ...obj,
                [r[1]]: {
                  [r[2]]: { [r[3]]: r[4] }
                }
              };
            }
            break;
          default:
            break;
        }
      }
    });
    return obj;
  } catch (error) {
    console.log('error: ', error);
    return {};
  }
};

const unProtectValueCell = row => {
  row.eachCell((cell, index) => {
    if (index === 4 || index === 5) {
      cell.fill = styles.fillCellForm;
      cell.protection = styles.protection.unlocked;
      if (_isEmpty(cell.value)) cell.fill = styles.fillEmptyCellForm;
    }
  });
  return row;
};
