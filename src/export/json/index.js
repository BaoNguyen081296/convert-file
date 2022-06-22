import fs from 'file-saver';
import { formatDate, TYPE } from 'utils/utils';
import _isEmpty from 'lodash/isEmpty';
const ExcelJS = require('exceljs');

export const exportFile = async ({ type = TYPE.TO_JSON, file }) => {
  if (type === TYPE.TO_JSON) handleExcelToJson(file);
  else {
    const reader = new FileReader();
    reader.addEventListener('load', e => {
      handleJsonToExcel(JSON.parse(reader.result));
    });
    reader.readAsText(file);
  }
};

const handleJsonToExcel = file => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'DeHR';
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  const worksheet = workbook.addWorksheet('JsonToExcel');
  transformDataToXLSX(file, worksheet);
  workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    fs.saveAs(blob, 'JsonToExcel-' + formatDate(new Date(), 'YYYYMMDDHHmm') + '.xlsx');
  });
};

const handleExcelToJson = async file => {
  const workbook = new ExcelJS.Workbook();
  let rows = (await workbook.xlsx.load(file)).getWorksheet()._rows;
  const data = transformDataToJson(rows);
  const blob = new Blob([JSON.stringify(data)], { type: 'text/plain;charset=utf-8' });
  return fs.saveAs(blob, 'convertedFile.json');
};

const transformDataToXLSX = (file, ws) => {
  try {
    if (typeof file === 'object') {
      const data = Object.keys(file);
      data.forEach(d => {
        if (typeof file[d] === 'object') {
          const keys = Object.keys(file[d]);
          keys.forEach(item => {
            if (typeof file[d][item] === 'object') {
              Object.keys(file[d][item]).forEach(i => {
                let row = [d, item, i, file[d][item][i]];
                ws.addRows([row]);
              });
            } else {
              let row = [d, item, file[d][item]];
              ws.addRows([row]);
            }
          });
        } else {
          ws.addRows([[d, file[d]]]);
        }
      });
    }
  } catch (error) {
    console.log('error: ', error);
    throw error;
  }
};

const transformDataToJson = rows => {
  try {
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
