// src/components/ExcelMerger.tsx

import React, { useState } from "react";
import * as XLSX from "xlsx";

interface ExcelData {
    fileName: string;
    data: any[][];
}

const ExcelMerger: React.FC = () => {
    const [mergedData, setMergedData] = useState<any[][]>([]);
    const [isLoading, setIsLoading] = useState(false);

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.files) {
            setIsLoading(true); // Set loading true while processing files
            readExcelFiles(e.target.files);
        }
    };

    const readExcelFiles = (files: FileList) => {
        const fileReaders: Promise<ExcelData>[] = [];
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const promise = new Promise<ExcelData>((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (event: ProgressEvent<FileReader>) => {
                    const data = event.target?.result;
                    if (data) {
                        const workbook = XLSX.read(data, {
                            type: file.name.split('.').pop() === "csv" ? "string" : "binary"
                        });
                        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                        const jsonData = XLSX.utils.sheet_to_json<any[]>(firstSheet, { header: 1, defval: "" });
                        resolve({ fileName: file.name, data: jsonData });
                    }
                };
                reader.onerror = (error) => reject(error);
                reader.readAsBinaryString(file);
            });
            fileReaders.push(promise);
        }

        Promise.all(fileReaders).then(dataArray => {
            const header = dataArray[0].data[14]; // 15th row as header (index 14)
            const data = dataArray.map((file) =>
                file.data.slice(15) // except first file starts from 15
            ).flat();
            setMergedData([header, ...data]);
            setIsLoading(false); // Reset loading state after processing
        });
    };

    const downloadMergedFile = () => {
        const ws = XLSX.utils.aoa_to_sheet(mergedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "MergedData");
        XLSX.writeFile(wb, "MergedData.xlsx");
    };

    return (
        <div className="max-w-xl mx-auto mt-10 p-5 border rounded-lg shadow-md">
            <h2 className="text-2xl font-bold mb-4">Excel Merger</h2>
            <input
                type="file"
                accept=".xlsx, .xls, .csv"
                multiple
                onChange={handleFileChange}
                className="mb-4 block w-full text-sm text-gray-500
                   file:mr-4 file:py-2 file:px-4
                   file:rounded-full file:border-0
                   file:text-sm file:font-semibold
                   file:bg-blue-50 file:text-blue-700
                   hover:file:bg-blue-100"
            />
            {isLoading && <div>Loading...</div>}
            {mergedData.length > 0 && (
                <div className="mt-5">
                    <button
                        onClick={downloadMergedFile}
                        className="bg-green-500 text-white px-4 py-2 rounded-lg hover:bg-green-600 transition"
                    >
                        Download Merged File
                    </button>
                </div>
            )}
        </div>
    );
};

export default ExcelMerger;
