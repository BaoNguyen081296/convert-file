import fs from 'file-saver';
const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();

const exportJSONFile = (data = null) => {
  let user = {
    name: 'Anonystick',
    emai: 'anonystick@gmail.com',
    age: 37,
    gender: 'Male',
    profession: 'Software Developer'
  };

  // convert JSON object to a string

  // write file to disk
  const blob = new Blob([JSON.stringify(data || user)], { type: 'text/plain;charset=utf-8' });
  return fs.saveAs(blob, 'convertedFile.json');
};

export const exportFile = async e => {
  let data = {};
  let rows = (await workbook.xlsx.load(e.originFileObj)).getWorksheet()._rows;
  rows.forEach(i => {
    const { values } = i;
    data = {
      ...data,
      [values[1]]: { ...(data[values[1]] ? data[values[1]] : {}), [values[2]]: values[3] }
    };
  });
  exportJSONFile(data);
};
