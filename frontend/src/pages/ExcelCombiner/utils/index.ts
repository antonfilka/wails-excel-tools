import {
  OpenExcelFile,
  GetFileSheetsNames,
  CombineExcelFiles,
} from "../../../../wailsjs/go/main/App";

export const openExcelFile = async () => {
  return await OpenExcelFile();
};
export const getFileSheets = async (file: string) => {
  return await GetFileSheetsNames(file);
};
export const combineFiles = async (
  file1: string,
  file2: string,
  sheet1: string,
  sheet2: string,
) => {
  await CombineExcelFiles(file1, file2, sheet1, sheet2);
};
