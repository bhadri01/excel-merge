// src/components/ExcelMerger.tsx

import React, { useState } from "react";
import * as XLSX from "xlsx";

interface ExcelData {
    fileName: string;
    data: any[][];
    isOpen: boolean;
}

const ExcelMerger: React.FC = () => {
    const [excelData, setExcelData] = useState<ExcelData[]>([]);
    const [selectedRows, setSelectedRows] = useState<{ [key: string]: boolean[] }>({});
    const [mergedData, setMergedData] = useState<any[][]>([]);

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.files) {
            readExcelFiles(e.target.files);
        }
    };

    const readExcelFiles = (files: FileList) => {
        const fileReaders: Promise<void>[] = [];
        setExcelData([]);
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const promise = new Promise<void>((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (event: ProgressEvent<FileReader>) => {
                    const data = event.target?.result;
                    if (data) {
                        const workbook = XLSX.read(data, {
                            type: file.name.split('.').pop() === "csv" ? "string" : "binary"
                        });
                        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                        const jsonData = XLSX.utils.sheet_to_json<any[]>(firstSheet, { header: 1 });
                        setExcelData((prev) => [
                            ...prev,
                            { fileName: file.name, data: jsonData, isOpen: false },
                        ]);
                        setSelectedRows((prev) => ({
                            ...prev,
                            [file.name]: Array(jsonData.length).fill(false),
                        }));
                    }
                    resolve();
                };
                reader.onerror = (error) => reject(error);
                reader.readAsBinaryString(file);
            });
            fileReaders.push(promise);
        }
    };

    const toggleRowSelection = (fileName: string, index: number) => {
        setSelectedRows((prev) => ({
            ...prev,
            [fileName]: prev[fileName].map((selected, i) =>
                i === index ? !selected : selected
            ),
        }));
    };

    const toggleAllRowsSelection = (fileName: string, isSelected: boolean) => {
        setSelectedRows(prev => ({
            ...prev,
            [fileName]: prev[fileName].map(() => isSelected)
        }));
    };

    const toggleAccordion = (fileName: string) => {
        setExcelData(prev => prev.map(sheet => sheet.fileName === fileName ? { ...sheet, isOpen: !sheet.isOpen } : sheet));
    };

    const mergeSelectedData = () => {
        let merged: any[][] = [];

        excelData.forEach(({ fileName, data }) => {
            const fileSelectedRows = selectedRows[fileName];
            data.forEach((row, rowIndex) => {
                if (fileSelectedRows[rowIndex]) {
                    merged.push(row);
                }
            });
        });

        setMergedData(merged);
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

            {excelData.map(({ fileName, data, isOpen }) => (
                <div key={fileName} className="mb-4">
                    <h3 className="text-lg font-semibold mb-2 cursor-pointer" onClick={() => toggleAccordion(fileName)}>
                        {fileName} {isOpen ? '-' : '+'}
                    </h3>
                    {isOpen && (
                        <div className="overflow-x-auto">
                            <button onClick={() => toggleAllRowsSelection(fileName, true)} className="bg-blue-200 mr-2 p-2 rounded text-sm">Select All</button>
                            <button onClick={() => toggleAllRowsSelection(fileName, false)} className="bg-red-200 p-2 rounded text-sm">Deselect All</button>
                            <table className="min-w-full divide-y divide-gray-200 mt-2">
                                <tbody className="bg-white divide-y divide-gray-200">
                                    {data.map((row, rowIndex) => (
                                        <tr key={rowIndex}>
                                            <td className="px-4 py-2">
                                                <input
                                                    type="checkbox"
                                                    checked={selectedRows[fileName][rowIndex]}
                                                    onChange={() => toggleRowSelection(fileName, rowIndex)}
                                                />
                                            </td>
                                            {row.map((cell, cellIndex) => (
                                                <td
                                                    key={cellIndex}
                                                    className="px-6 py-4 whitespace-nowrap text-sm text-gray-500"
                                                >
                                                    {cell}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            ))}

            <button
                onClick={mergeSelectedData}
                className="bg-blue-500 text-white px-4 py-2 rounded-lg hover:bg-blue-600 transition"
            >
                Merge Selected Data
            </button>
            {mergedData.length > 0 && (
                <div className="mt-5">
                    <h3 className="text-xl font-bold mb-2">Merged Data Preview:</h3>
                    <div className="overflow-x-auto">
                        <table className="min-w-full divide-y divide-gray-200">
                            <thead className="bg-gray-50">
                                <tr>
                                    {mergedData[0]?.map((header, index) => (
                                        <th
                                            key={index}
                                            className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider"
                                        >
                                            {header}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-gray-200">
                                {mergedData.map((row, rowIndex) => (
                                    <tr key={rowIndex}>
                                        {row.map((cell, cellIndex) => (
                                            <td
                                                key={cellIndex}
                                                className="px-6 py-4 whitespace-nowrap text-sm text-gray-500"
                                            >
                                                {cell}
                                            </td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                        <button
                            onClick={downloadMergedFile}
                            className="bg-green-500 text-white px-4 py-2 rounded-lg hover:bg-green-600 transition mt-2"
                        >
                            Download Merged File
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
};

export default ExcelMerger;
