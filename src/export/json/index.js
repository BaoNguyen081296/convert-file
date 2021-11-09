import fs from 'file-saver';
import { formatDate, TYPE } from 'utils/utils';
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
  console.log('file: ', file);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'DeHR';
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  const worksheet = workbook.addWorksheet('JsonToExcel');
  handleTransformData(file, worksheet);
  workbook.xlsx.writeBuffer().then(data => {
    // const blob = new Blob([data], {
    //   type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    // });
    // fs.saveAs(blob, 'JsonToExcel-' + formatDate(new Date(), 'YYYYMMDDHHmm') + '.xlsx');
  });
};

const handleExcelToJson = async file => {
  const workbook = new ExcelJS.Workbook();
  let data = {};
  let rows = (await workbook.xlsx.load(file)).getWorksheet()._rows;
  rows.forEach(i => {
    const { values } = i;
    data = {
      ...data,
      [values[1]]: { ...(data[values[1]] ? data[values[1]] : {}), [values[2]]: values[3] }
    };
  });
  const blob = new Blob([JSON.stringify(data)], { type: 'text/plain;charset=utf-8' });
  return fs.saveAs(blob, 'convertedFile.json');
};

const handleTransformData = (obj, ws) => {
  if (typeof obj === 'object') {
    const data = Object.keys(obj);
    let rows = [];
    data.forEach(i => {
      console.log('obj: ', obj);
      if (typeof obj[i] === 'object') {
        handleTransformData(obj[i]);
      } else {
        rows = [i];
      }
      // ws.addRows(rows);
    });
    console.log('data: ', data);
  }
};
