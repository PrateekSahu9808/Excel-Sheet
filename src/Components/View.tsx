import React, { useRef, useEffect, useState } from "react";
import { useParams } from "react-router-dom";
import { Button } from "@mui/material";
import Spreadsheet from "x-data-spreadsheet";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

const View: React.FC = () => {
  const { key } = useParams<{ key: string }>();
  const [data, setData] = useState<Spreadsheet | null| []>([]);
  const spreadsheetRef = useRef<Spreadsheet | null>(null);
  console.log(data);
  useEffect(() => {
    const loadDataByKey = () => {
      if (key === undefined) return;
      const storedData = localStorage.getItem(key);

      if (storedData) {
        setData(JSON.parse(storedData));
      } else {
        setData(null);
      }
    };
    loadDataByKey();
  }, [key]);

  useEffect(() => {
    const createNewSpreadsheet = () => {
      const newSpreadsheet = new Spreadsheet("#spreadsheet", {
        mode: "edit",
        showToolbar: true,
        showGrid: true,
        // data: [],
        row: {
          len: 100,
          height: 25,
        },
        col: {
          len: 26,
          width: 100,
          indexWidth: 60,
          minWidth: 60,
        },
      });
      spreadsheetRef.current = newSpreadsheet;
    };
    createNewSpreadsheet();
  }, []);

  const handleLoadData = () => {
    if (data && spreadsheetRef.current) {
      spreadsheetRef.current.loadData(data);
    }
  };

  const handleExport = () => {
    if (spreadsheetRef.current) {
      const newData = spreadsheetRef.current.getData();
     
      exportSheet(newData, `${key}.xlsx`);
    } else {
      console.error("Spreadsheet instance is not available.");
    }
  };

  useEffect(() => {
    handleLoadData();
  }, [data]);

  const handleSaveData = () => {
    if (spreadsheetRef.current) {
      if (key === undefined) return;
      const newData = spreadsheetRef.current.getData();
      localStorage.setItem(key, JSON.stringify(newData));
      alert(`Data is successfully saved on the file ${key}`);
    } else {
      console.error("Spreadsheet instance is not available.");
    }
  };

  const exportSheet = (sdata:any, filename:any) => {
    const workbook = new ExcelJS.Workbook();
    

    sdata.forEach((sheet:any) => {
      const { name, rows, styles } = sheet;
      const worksheet = workbook.addWorksheet(name);

      for (let colIndex = 1; colIndex <= 26; colIndex++) {
        worksheet.getColumn(colIndex).width = 15;
      }

      if (styles && styles.length > 0) {
        styles.forEach((style:any, styleIndex:any) => {
          const rowNumber = styleIndex + 1;

          if (style.cells) {
            Object.keys(style.cells).forEach((colIndex) => {
              const cell = worksheet.getCell(
                `${getColumnName(+colIndex + 1)}${rowNumber}`
              );
              applyStyles(cell, style.cells[colIndex]);
            });
          }
        });
      }

      if (rows && Object.keys(rows).length > 0) {
        const rowIndices = Object.keys(rows)
          .map(Number)
          .sort((a, b) => a - b);

        rowIndices.forEach((rowIndex) => {
          const rowData = rows[rowIndex];
          if (rowData && rowData.cells) {
            const newRow = worksheet.getRow(rowIndex + 1);

            Object.keys(rowData.cells).forEach((colIndex) => {
              const cell = newRow.getCell(+colIndex + 1);
              const cellData = rowData.cells[colIndex]?.text;

              if (cellData !== undefined) {
                cell.value = cellData;

                if (rowData.cells[colIndex].style !== undefined) {
                  const styleIndex = rowData.cells[colIndex].style;
                  const cellStyle = styles[styleIndex];
                  applyStyles(cell, cellStyle);
                }
              }
            });
          }
        });
      }

      if (sheet.merges) {
        sheet.merges.forEach((range:any) => {
          worksheet.mergeCells(range);
        });
      }
    });

    workbook.xlsx
      .writeBuffer()
      .then((buffer) => {
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        saveAs(blob, filename);
        console.log("File saved successfully.");
      })
      .catch((error) => {
        console.error("Error saving file:", error);
      });
  };

  const getColumnName = (colNumber:any) => {
    let dividend= colNumber + 1;
    let columnName = "";
    let modulo
   
    while (dividend > 0) {
      modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      // dividend = parseInt((dividend - modulo) / 26, 10);
      dividend = parseInt(String((dividend - modulo) / 26), 10);
    }

    return columnName;
  };

  const applyStyles = (cell:any, style:any) => {
    if (style) {
      if (style.font) {
        cell.font = cell.font || {};
        if (style.font.bold) cell.font.bold = style.font.bold;
        if (style.font.italic) cell.font.italic = style.font.italic;
        if (style.font.size) cell.font.size = style.font.size;
        if (style.font.color) cell.font.color = { rgb: style.font.color };
        if (style.font.strike) cell.font.strike = style.font.strike;
        if (style.font.name) cell.font.name = style.font.name;
      }

      if (style.hasOwnProperty("underline")) {
        console.log(style.underline);
        cell.font = cell.font || {};
        cell.font.underline = !!style.underline;
      }

      if (style.bgcolor) {
        cell.fill = cell.fill || {};
        cell.fill.type = "pattern";
        cell.fill.pattern = "solid";
        let rgbColor = style.bgcolor;
        let argb = rgbColor.slice(1);
        cell.fill.fgColor = { argb: argb };
      }

      if (style.color) {
        let rgbColor = style.color;
        let argb = rgbColor.slice(1);
        cell.font = cell.font || {};
        cell.font.color = { argb: argb };
      }

      if (style.border) {
        cell.border = cell.border || {};
        if (style.border.top) {
          cell.border.top = {
            style: "thin",
          };
        }
        if (style.border.bottom) {
          cell.border.bottom = {
            style: "thin",
          };
        }
        if (style.border.left) {
          cell.border.left = {
            style: "thin",
          };
        }
        if (style.border.right) {
          cell.border.right = {
            style: "thin",
          };
        }
      }
    }
  };


  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files && event.target.files[0];

    if (file) {
      const reader = new FileReader();

      reader.onload = async () => {
        try {
          const workbook:any = new ExcelJS.Workbook();
          await workbook.xlsx.load(file);

          // const importedSheets:[] = [];
          const importedSheets: { name: string; rows: any }[] = [];

          workbook.eachSheet((worksheet:any) => {
            // const rows = {};
             const rows: { [key: number]: any } = {};

            worksheet.eachRow((row:any, rowNumber:any) => {
              // const cells = {};
                 const cells: { [key: number]: any } = {};
              row.eachCell((cell:any, colNumber:any) => {
                const adjustedColNumber = colNumber - 1;
                
                
                cells[adjustedColNumber] = {
                  text: cell.text || "",
                  font: cell.style.font || {},
                  alignment: cell.alignment || {},
                  border: cell.border || {},
                  fill: cell.fill || { type: "pattern", pattern: "solid" },
                };
              });

              const adjustedRowNumber = rowNumber - 1;

              rows[adjustedRowNumber] = {
                cells,
                __rowNum__: adjustedRowNumber,
              };
            });

            importedSheets.push({ name: worksheet.name, rows });
          });
          const existingData = data || [];

          // const mergedData = existingData.concat(importedSheets);
          const mergedData:any = (existingData as any[]).concat(importedSheets);

          setData(mergedData);

          if (spreadsheetRef.current) {
            spreadsheetRef.current.loadData(mergedData);
          }
        } catch (error) {
          console.error("Error handling file change:", error);
        }
      };

      reader.readAsBinaryString(file);
    }
  };

  return (
    <div>
      <div className="flex justify-end gap-2 mr-4 ">
        <Button
          variant="contained"
          className="z-50"
          onClick={() => handleSaveData()}
        >
          Save Data
        </Button>

        <Button
          onClick={() => handleExport()}
          variant="contained"
          className="z-50 "
        >
          Export Data
        </Button>
        {/* <input type="file" onChange={handleFileUpload} className="z-50" /> */}

        <label htmlFor="file-upload">
          <Button variant="contained" component="span" className="z-50">
            Import Excel
          </Button>
        </label>
        <input
          type="file"
          id="file-upload"
          accept=".xlsx, .xls"
          style={{ display: "none" }}
          onChange={handleFileUpload}
          className="z-50"
        />
      </div>
      <div id="spreadsheet" className="absolute top-0  "></div>
    </div>
  );
};

export default View;
