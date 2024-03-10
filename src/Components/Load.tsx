
import React, { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Link } from "react-router-dom";
import {
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Button,
  Box,
  Typography,
} from "@mui/material";
const Load: React.FC = () => {
  const [loadedData, setLoadedData] = useState<string[]>([]);
  const navigate = useNavigate();
  useEffect(() => {
    const getAllLocalStorageKeys = () => {
      const keys: string[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key) keys.push(key);
      }
      return keys;
    };
    const keysFromLocalStorage = getAllLocalStorageKeys();
    setLoadedData(keysFromLocalStorage);
  }, []);
  return (
    <div style={{ textAlign: "center" }}>
      <h2 className="font-bold text-3xl pt-4">All Files are here !!!</h2>
      <Box sx={{ paddingTop: "2rem" }}>
        <TableContainer component={Paper}>
          <Table>
            <TableHead>
              <TableRow>
                <TableCell className="font-bold">
                  <Typography variant="h6" sx={{ fontWeight: "600" }}>
                    File Name
                  </Typography>
                </TableCell>
                <TableCell>
                  <Typography variant="h6" sx={{ fontWeight: "600" }}>
                    Action
                  </Typography>
                </TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {loadedData.map((key, index) => (
                <TableRow key={index}>
                  <TableCell>{key}</TableCell>
                  <TableCell>
                    <Link to={`/view/${key}`}>
                      <Button variant="contained">View</Button>
                    </Link>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      </Box>
      <div className="flex justify-end p-4">
        <Button variant="contained" onClick={() => navigate("/")}>
          Back to home
        </Button>
      </div>
    </div>
  );
};
export default Load;
