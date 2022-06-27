import moment from 'moment';

export const TYPE = {
  TO_XLSX: 'toXLSX',
  TO_JSON: 'toJson',
  XLSX: 'XLSX',
  JSON: 'JSON'
};

export const formatDate = (inputDate, formatString = 'DD/MM/YYYY') => {
  return inputDate && moment(inputDate).format(formatString);
};

export const EXCEL_PASSWORD = 'DEHR_DEV';
