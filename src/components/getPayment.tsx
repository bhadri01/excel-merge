import React, { useState } from 'react';
import * as XLSX from 'xlsx';

interface PaymentSummary {
    sno: number;
    challanDate: string;
    noOfChallans: number;
    noOfPayments: number;
    count100: number;
    total100: number;
    count1000: number;
    total1000: number;
    grandTotal: number;
}

const ExcelMerger: React.FC = () => {
    const [paymentSummaries, setPaymentSummaries] = useState<PaymentSummary[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(false);

    const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        const files = event.target.files;
        if (files && files[0]) {
            setIsLoading(true);
            readExcelFile(files[0]);
        }
    };

    const readExcelFile = (file: File) => {
        const reader = new FileReader();
        reader.onload = (e: ProgressEvent<FileReader>) => {
            const data = e.target?.result;
            if (typeof data === 'string') {
                const workbook = XLSX.read(data, { type: 'binary' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                processPaymentData(jsonData);
            }
        };
        reader.readAsBinaryString(file);
    };

    const excelSerialDateToDate = (serial: number): Date => {
        const excelStartDate = new Date(Date.UTC(1899, 11, 31));
        const days = Math.floor(serial);
        const millisecondsInADay = 86400000;
        const adjustedMilliseconds = (serial - days) * millisecondsInADay; // time part

        return new Date(excelStartDate.getTime() + days * millisecondsInADay + adjustedMilliseconds);
    };

    const formatDate = (date: Date): string => {
        const d = date.getDate();
        const m = date.getMonth() + 1; // Months are zero indexed
        const y = date.getFullYear();

        return `${d.toString().padStart(2, '0')}/${m.toString().padStart(2, '0')}/${y}`;
    };

    const processPaymentData = (data: any[]) => {
        const summary: PaymentSummary[] = data.slice(1).reduce((acc: PaymentSummary[], row: any[]) => {
            const challanDate = excelSerialDateToDate(parseFloat(row[10]));
            const paymentDate = excelSerialDateToDate(parseFloat(row[11]));
            const challanDateString = formatDate(challanDate);
            const paymentDateString = formatDate(paymentDate);
            const amount = parseFloat(row[13]);
            const currentDate = new Date();

            if (challanDate >= new Date('2021/12/01') && challanDate <= currentDate) {
                const existing = acc.find(item => item.challanDate === challanDateString) || {
                    sno: acc.length + 1,
                    challanDate: challanDateString,
                    noOfChallans: 0,
                    noOfPayments: 0,
                    count100: 0,
                    total100: 0,
                    count1000: 0,
                    total1000: 0,
                    grandTotal: 0
                };

                existing.noOfChallans++;
                if (paymentDateString) {
                    existing.noOfPayments++;
                }

                if (amount === 100) {
                    existing.count100++;
                    existing.total100 += amount;
                } else if (amount === 1000) {
                    existing.count1000++;
                    existing.total1000 += amount;
                }

                existing.grandTotal = existing.total100 + existing.total1000;

                if (!acc.find(item => item.challanDate === challanDateString)) {
                    acc.push(existing);
                }
            }
            return acc;
        }, []);

        setPaymentSummaries(summary);
        setIsLoading(false);
    };

    const downloadSummaryAsExcel = () => {
        const ws = XLSX.utils.json_to_sheet(paymentSummaries);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Summary");
        XLSX.writeFile(wb, "PaymentSummary.xlsx");
    };

    return (
        <div className="app-container">
            <h1>Excel Payment Processor</h1>
            <input type="file" onChange={handleFileChange} accept=".xlsx, .xls, .csv" />
            {isLoading && <p>Loading...</p>}
            {paymentSummaries.length > 0 && (
                <div>
                    <h2>Payment Summary</h2>
                    <button onClick={downloadSummaryAsExcel} className="bg-blue-500 text-white px-4 py-2 rounded-lg hover:bg-blue-700">
                        Download Summary
                    </button>
                </div>
            )}
        </div>
    );
};

export default ExcelMerger;
