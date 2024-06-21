const XLSX = require('xlsx');
const moment = require('moment');
const fs = require('fs');

const employeeSheetPath = 'C:/Users/ELCOT/Documents/excel_sheet/Book1.xlsx';
const timesheetSheetPath = 'C:/Users/ELCOT/Documents/excel_sheet/Book2.xlsx';

try {
   
    const employeeWorkbook = XLSX.readFile(employeeSheetPath);
    const employeeSheet = XLSX.utils.sheet_to_json(employeeWorkbook.Sheets[employeeWorkbook.SheetNames[0]]);

    const timesheetWorkbook = XLSX.readFile(timesheetSheetPath);
    const timesheetSheet = XLSX.utils.sheet_to_json(timesheetWorkbook.Sheets[timesheetWorkbook.SheetNames[0]]);

    const WORK_HOURS_PER_DAY = 9;
    const LOP_THRESHOLD = 3;  

    function calculateSalary(employee, logs) {
        let lopInstances = 0;

        logs.forEach(log => {
            const inTime = moment(log.InTime, 'HH:mm');
            const outTime = moment(log.OutTime, 'HH:mm');
            const workHours = outTime.diff(inTime, 'hours');
            
            if (workHours < WORK_HOURS_PER_DAY) {
                lopInstances++;
            }
        });

        // Calculate LOP deduction
        const lopDays = Math.floor(lopInstances / LOP_THRESHOLD) * 0.5;
        const lopAmount = lopDays * (employee.Salary / 30);

        // Calculate total salary after deducting LOP
        const totalSalary = employee.Salary - lopAmount;

        return {
            employeeId: employee.EmployeeId,
            employeeName: employee.Name,
            baseSalary: employee.Salary,
            lopInstances,
            lopAmount,
            totalSalary
        };
    }

    const results = employeeSheet.map(employee => {
        const employeeLogs = timesheetSheet.filter(log => log.EmployeeId === employee.EmployeeId);
        return calculateSalary(employee, employeeLogs);
    });

    fs.writeFileSync('./payroll_results.json', JSON.stringify(results, null, 2));

    console.log('Payroll calculated successfully. Results saved to payroll_results.json');
} catch (error) {
    console.error('An error occurred:', error.message);
}









