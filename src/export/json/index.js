/* eslint-disable no-loop-func */
import fs from 'file-saver';
import { formatDate, TYPE, EXCEL_PASSWORD, DESTINATION_FILE } from 'utils/utils';
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

export const exportFile = async ({ type = TYPE.TO_JSON, file, destination }) => {
  try {
    if (type === TYPE.TO_JSON) handleExcelToJson(file, destination);
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

const handleExcelToJson = async (file, destination) => {
  let rows = (await workbook.xlsx.load(file[0].originFileObj)).getWorksheet()._rows;
  rows.shift();
  const data = transformDataToJson(rows, destination);
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

const transformDataToJson = (rows, destination) => {
  try {
    let obj = {};
    const iKey = {
      first: 1,
      second: 2,
      third: 3
    };
    rows.forEach(i => {
      const { values: curArr } = i;
      if (!_isEmpty(curArr)) {
        let currentDestinationIndex = 0;
        if (destination === DESTINATION_FILE.TO_FIRST) {
          currentDestinationIndex = 4;
        } else {
          currentDestinationIndex = 5;
        }
        const { first, second, third } = iKey;
        const curIndexValue = curArr[currentDestinationIndex];
        if (_isEmpty(curArr[third])) {
          obj = {
            ...obj,
            [curArr[first]]: {
              ...(obj[curArr[first]]
                ? { ...obj[curArr[first]], [curArr[second]]: curIndexValue }
                : { [curArr[second]]: curIndexValue })
            }
          };
        } else {
          if (_isEmpty(obj[curArr[first]])) {
            obj = {
              ...obj,
              [curArr[first]]: { [curArr[second]]: { [curArr[third]]: curIndexValue } }
            };
          } else {
            if (_isEmpty(obj[curArr[first]][curArr[second]])) {
              obj = {
                ...obj,
                [curArr[first]]: {
                  ...obj[curArr[first]],
                  [curArr[second]]: { [curArr[third]]: curIndexValue }
                }
              };
            } else {
              obj = {
                ...obj,
                [curArr[first]]: {
                  ...obj[curArr[first]],
                  [curArr[second]]: {
                    ...obj[curArr[first]][curArr[second]],
                    [curArr[third]]: curIndexValue
                  }
                }
              };
            }
          }
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
