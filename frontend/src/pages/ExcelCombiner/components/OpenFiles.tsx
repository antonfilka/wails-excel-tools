import { useEffect, useState } from "react";
import { openExcelFile, getFileSheets, combineFiles } from "../utils";
import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectGroup,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Input } from "@/components/ui/input";

export const OpenFiles = () => {
  const [files, setFiles] = useState({
    file1: "",
    file2: "",
    file1Sheet: "",
    file2Sheet: "",
  });

  const [file1Sheets, setFile1Sheets] = useState<string[]>([]);
  const [file2Sheets, setFile2Sheets] = useState<string[]>([]);

  const handleOpenClick = async (fileNumber: number) => {
    const file = await openExcelFile();
    if (fileNumber === 1) {
      setFiles((prev) => ({ ...prev, file1: file }));
    } else if (fileNumber === 2) {
      setFiles((prev) => ({ ...prev, file2: file }));
    }
  };

  useEffect(() => {
    const updateFilesSheets = async () => {
      if (files.file1) {
        const sheets = await getFileSheets(files.file1);
        setFile1Sheets(sheets);
      }
      if (files.file2) {
        const sheets = await getFileSheets(files.file2);
        setFile2Sheets(sheets);
      }
    };

    updateFilesSheets();
  }, [files.file1, files.file2]);

  const handleCombineClick = async () => {
    await combineFiles(
      files.file1,
      files.file2,
      files.file1Sheet,
      files.file2Sheet,
    );
  };

  const isButtonDisabled =
    !files.file1 || !files.file2 || !files.file1Sheet || !files.file2Sheet;

  return (
    <div className="w-full h-full flex flex-col gap-3">
      <div className="w-full h-full flex items-start justify-between">
        <div className="w-[48%] flex flex-col items-center justify-start gap-3">
          <Button className="w-full" onClick={() => handleOpenClick(1)}>
            Выбрать общий XLSX файл
          </Button>
          <Input
            className="w-full"
            value={files.file1}
            readOnly
            placeholder="File"
          />
          {file1Sheets.length > 0 && (
            <div className="w-full flex flex-col gap-2">
              <Select
                onValueChange={(val) => {
                  setFiles((prev) => ({ ...prev, file1Sheet: val }));
                }}
                defaultValue={"Не выбрано"}
                value={files.file1Sheet}
              >
                <SelectTrigger className="w-full">
                  <SelectValue placeholder="Выберите страницу"></SelectValue>
                </SelectTrigger>
                <SelectContent>
                  <SelectGroup>
                    {file1Sheets.map((sheet) => (
                      <SelectItem key={sheet} value={sheet}>
                        {sheet}
                      </SelectItem>
                    ))}
                  </SelectGroup>
                </SelectContent>
              </Select>
            </div>
          )}
        </div>

        <div className="w-[48%] flex flex-col items-center justify-start gap-3">
          <Button className="w-full" onClick={() => handleOpenClick(2)}>
            Выбрать 1C XLSX файл
          </Button>
          <Input
            className="w-full"
            value={files.file2}
            readOnly
            placeholder="File"
          />
          {file2Sheets.length > 0 && (
            <div className="w-full flex flex-col gap-2">
              <Select
                onValueChange={(val) =>
                  setFiles((prev) => ({ ...prev, file2Sheet: val }))
                }
                defaultValue={"Не выбрано"}
                value={files.file2Sheet}
              >
                <SelectTrigger className="w-full">
                  <SelectValue placeholder="Выберите страницу"></SelectValue>
                </SelectTrigger>
                <SelectContent>
                  <SelectGroup>
                    {file2Sheets.map((sheet) => (
                      <SelectItem key={sheet} value={sheet}>
                        {sheet}
                      </SelectItem>
                    ))}
                  </SelectGroup>
                </SelectContent>
              </Select>
            </div>
          )}
        </div>
      </div>
      <Button
        className="w-full"
        onClick={handleCombineClick}
        disabled={isButtonDisabled}
      >
        Старт
      </Button>
    </div>
  );
};
