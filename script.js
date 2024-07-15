function processFiles() {
    const currentMonthFileInput = document.getElementById('currentMonthFileInput');
    const previousMonthFileInput = document.getElementById('previousMonthFileInput');
    const otDaysInput = document.getElementById('otDaysInput').value;
    const fixedSalaryInput = document.getElementById('fixedSalaryInput').value;

    if (!currentMonthFileInput.files.length) {
        alert('Please select the current month file!');
        return;
    }

    if (!previousMonthFileInput.files.length && !otDaysInput) {
        alert('Please select the previous month file or enter the number of OT days!');
        return;
    }

    const currentMonthFile = currentMonthFileInput.files[0];
    const previousMonthFile = previousMonthFileInput.files.length ? previousMonthFileInput.files[0] : null;

    const currentMonthReader = new FileReader();
    const previousMonthReader = new FileReader();

    let currentMonthData, previousMonthData;

    currentMonthReader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        currentMonthData = XLSX.read(data, { type: 'array' });
        if (previousMonthData || otDaysInput) processCalculations(currentMonthData, previousMonthData, fixedSalaryInput, otDaysInput);
    };

    if (previousMonthFile) {
        previousMonthReader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            previousMonthData = XLSX.read(data, { type: 'array' });
            if (currentMonthData) processCalculations(currentMonthData, previousMonthData, fixedSalaryInput, otDaysInput);
        };
        previousMonthReader.readAsArrayBuffer(previousMonthFile);
    }

    currentMonthReader.readAsArrayBuffer(currentMonthFile);
}

function processCalculations(currentMonthData, previousMonthData, fixedSalary, otDaysInput) {
    const sheetCurrent = currentMonthData.Sheets[currentMonthData.SheetNames[0]];
    const sheetPrevious = previousMonthData ? previousMonthData.Sheets[previousMonthData.SheetNames[0]] : null;

    let currentMonthDate, currentMonthDays, currentMonthHolidays, currentMonthLeaves;
    let previousMonthDate, previousMonthDays, previousMonthHolidays;

    currentMonthDate = getDateFromSheet(sheetCurrent, 'B10');
    if (sheetPrevious) {
        previousMonthDate = getDateFromSheet(sheetPrevious, 'B10');
    } else {
        previousMonthDate = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth() - 1, 1);
    }

        currentMonthDays = new Date(currentMonthDate.getFullYear(), currentMonthDate.getMonth() + 1, 0).getDate();
    previousMonthDays = new Date(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1, 0).getDate();

    currentMonthHolidays = countHolidaysInMonth(currentMonthDate);
    previousMonthHolidays = countHolidaysInMonth(previousMonthDate);

    currentMonthLeaves = getLeavesFromSheet(sheetCurrent, 'G5');

    const daysWorkedCurrentMonth = currentMonthDays - currentMonthLeaves;

    let proratedFixedSalary = Number(fixedSalary);
    if (daysWorkedCurrentMonth < currentMonthDays) {
        proratedFixedSalary = (fixedSalary / currentMonthDays) * daysWorkedCurrentMonth;
    }

    let previousMonthOTDays = 0;
    if (sheetPrevious) {
        const totalPreviousMonthMinutes = getTotalMinutesFromSheet(sheetPrevious);
        const previousMonthOTHours = totalPreviousMonthMinutes / 60 - previousMonthDays * 8;
        previousMonthOTDays = previousMonthOTHours > 0 ? Math.floor(previousMonthOTHours / 8) : 0;
    } else {
        previousMonthOTDays = Number(otDaysInput);
    }

    const amountFor1OTPrev = 1500 / (previousMonthDays - previousMonthHolidays) * 1.25;
    const previousMonthOTSalary = amountFor1OTPrev * previousMonthOTDays;

    const currentMonthHolidaySalary = 1500 / (currentMonthDays - currentMonthHolidays) * 1.25 * currentMonthHolidays;

    const totalSalary = proratedFixedSalary + currentMonthHolidaySalary + previousMonthOTSalary;

    const monthNameCurrent = currentMonthDate.toLocaleString('default', { month: 'long' });
    const monthNamePrevious = previousMonthDate.toLocaleString('default', { month: 'long' });

    const currentMonthInfo = `Current Month: ${monthNameCurrent}, Days: ${currentMonthDays}, Holidays: ${currentMonthHolidays}, Leaves Taken: ${currentMonthLeaves}, Days Worked: ${daysWorkedCurrentMonth}`;
    const previousMonthInfo = sheetPrevious ? `Previous Month: ${monthNamePrevious}, Days: ${previousMonthDays}, Holidays: ${previousMonthHolidays}` : 'Previous Month: OT days entered manually';
    const overtimeDays = `Overtime Days in Previous Month: ${previousMonthOTDays}`;
    const otSalary = `Previous Month OT Salary: ${previousMonthOTSalary.toFixed(2)}`;
    const holidaySalary = `Current Month Holiday Salary: ${currentMonthHolidaySalary.toFixed(2)}`;
    const totalSalaryText = `Total Salary: ${totalSalary.toFixed(2)}`;

    localStorage.setItem("currentMonthInfo", currentMonthInfo);
    localStorage.setItem("previousMonthInfo", previousMonthInfo);
    localStorage.setItem("overtimeDays", overtimeDays);
    localStorage.setItem("otSalary", otSalary);
    localStorage.setItem("holidaySalary", holidaySalary);
    localStorage.setItem("totalSalary", totalSalaryText);

    window.location.href = "results.html";
}

function getDateFromSheet(sheet, cellAddress) {
    const cell = sheet[cellAddress];
    if (cell && cell.v) {
        if (typeof cell.v === 'string') {
            const dateParts = cell.v.split(/[- /]/);
            if (dateParts.length === 3) {
                const day = parseInt(dateParts[0], 10);
                const month = parseInt(dateParts[1], 10) - 1;
                const year = parseInt(dateParts[2], 10);
                return new Date(year, month, day);
            }
        } else if (typeof cell.v === 'number') {
            const parsedDate = XLSX.SSF.parse_date_code(cell.v);
            return new Date(parsedDate.y, parsedDate.m - 1, parsedDate.d);
        }
    }
    return null;
}

function getTotalMinutesFromSheet(sheet) {
    let totalMinutes = 0;
    for (let row = 10; row <= 50; row++) {
        const cellAddress = `E${row}`;
        const cell = sheet[cellAddress];
        if (cell && cell.v) {
            let timeValue = cell.v;
            if (typeof timeValue === 'number') {
                const parsedTime = XLSX.SSF.parse_date_code(timeValue);
                timeValue = `${parsedTime.H}:${parsedTime.M}:${parsedTime.S}`;
            } else if (typeof timeValue === 'string') {
                const timeParts = timeValue.split(':');
                if (timeParts.length === 3) {
                    const hours = parseInt(timeParts[0], 10);
                    const minutes = parseInt(timeParts[1], 10);
                    totalMinutes += hours * 60 + minutes;
                }
            }
        }
    }
    return totalMinutes;
}

function countHolidaysInMonth(date) {
    const specificHolidays = [
        new Date(date.getFullYear(), 1, 13), // 13-Feb-24
        new Date(date.getFullYear(), 3, 10), // 10-Apr-24
        new Date(date.getFullYear(), 3, 11), // 11-Apr-24
        new Date(date.getFullYear(), 3, 13), // 13-Apr-24
        new Date(date.getFullYear(), 5, 16), // 16-Jun-24
        new Date(date.getFullYear(), 5, 17), // 17-Jun-24
        new Date(date.getFullYear(), 5, 18), // 18-Jun-24
        new Date(date.getFullYear(), 11, 18) // 18-Dec-24
    ];
    
    let holidayCount = 0;
    
    for (const holiday of specificHolidays) {
        if (holiday.getMonth() === date.getMonth()) {
            holidayCount++;
        }
    }

    const day = new Date(date.getFullYear(), date.getMonth(), 1);
    while (day.getMonth() === date.getMonth()) {
        if (day.getDay() === 5) { // Friday
            holidayCount++;
        }
        day.setDate(day.getDate() + 1);
    }

    return holidayCount;
}

function getLeavesFromSheet(sheet, cellAddress) {
    const cell = sheet[cellAddress];
    return cell && cell.v ? parseInt(cell.v, 10) : 0;
}
function disableOTDaysInput() {
    const otDaysInput = document.getElementById('otDaysInput');
    if (document.getElementById('previousMonthFileInput').files.length > 0) {
        otDaysInput.disabled = true;
        otDaysInput.value = '';
    } else {
        otDaysInput.disabled = false;
    }
}

function disablePreviousMonthFileInput() {
    const previousMonthFileInput = document.getElementById('previousMonthFileInput');
    if (document.getElementById('otDaysInput').value) {
        previousMonthFileInput.disabled = true;
        previousMonthFileInput.value = '';
    } else {
        previousMonthFileInput.disabled = false;
    }
}

