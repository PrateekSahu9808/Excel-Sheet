import React, { useEffect, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button, Input } from "@mui/material";
import Spreadsheet from "x-data-spreadsheet";

const Sheet: React.FC = () => {
 
  const [sheetName, setSheetName] = useState("");
  const el = useRef<HTMLDivElement>(null);
  const spreadsheetRef = useRef<Spreadsheet | null>(null);
   const navigate = useNavigate();
  const handleSendData = () => {
    if (sheetName && spreadsheetRef.current) {
      const key = sheetName;
      const data = spreadsheetRef.current.getData();
      localStorage.setItem(key, JSON.stringify(data));
      setSheetName("");

      const container = el.current;
      if (container) {
        container.innerHTML = "";
      }

      spreadsheetRef.current = new Spreadsheet(el.current!, {
        view: {
          height: () => document.documentElement.clientHeight,
          width: () => document.documentElement.clientWidth,
        },
        // data: {},
        showToolbar: true,
        showGrid: true,
      });
    }
    alert("sheet is created");
  };
  const handleSheetNameChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    setSheetName(event.target.value);
  };
  useEffect(() => {
    // Initial creation of the spreadsheet instance
    if (!spreadsheetRef.current) {
      spreadsheetRef.current = new Spreadsheet(el.current!, {
        view: {
          height: () => document.documentElement.clientHeight,
          width: () => document.documentElement.clientWidth,
        },
        // data: {},
        showToolbar: true,
        showGrid: true,
      });
    }
  }, []);
  return (
    <>
      <div>
        <div className="flex justify-between items-end gap-3 m-4">
          <div className="flex gap-3 ml-8">
            <Button variant="contained" onClick={() => navigate("/")}>
              Back to home
            </Button>
            <Button variant="contained" onClick={() => navigate("/load")}>
              All Sheets
            </Button>
          </div>
          <div className="flex items-end gap-4 mr-14">
            <Input
              type="text"
              placeholder="Enter sheet name"
              value={sheetName}
              onChange={handleSheetNameChange}
            />
            <Button
              variant="contained"
              onClick={handleSendData}
              disabled={!sheetName}
            >
              Save Data
            </Button>
          </div>
        </div>
        <div ref={el} className="mt-2"></div>
      </div>
    </>
  );
};
export default Sheet;
