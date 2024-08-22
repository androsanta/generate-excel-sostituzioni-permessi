import ExcelJS from 'exceljs';
import {addMonths, format, isBefore, isSameMonth, startOfMonth, setDefaultOptions, addDays, isSaturday, isSunday} from 'date-fns';
import {it} from 'date-fns/locale/index.js'
import {saveAs} from 'file-saver';


export async function _generateExcel(teachers, yearInput) {

    setDefaultOptions({locale: it})

    const workbook = new ExcelJS.Workbook();

    const year = yearInput || new Date().getFullYear();
    const startMonth = new Date(year, 8);
    const endMonth = new Date(year + 1, 5);

    const hasComma = teachers.includes(',');
    const hasNewLine = teachers.includes('\n');

    if (hasComma && hasNewLine) {
        throw new Error('Only one separator allowed');
    }

    const separator = hasComma ? ',' : '\n';
    const teachersList = teachers.split(separator).map(t => t.trim()).sort();

    console.log('year', year, 'startMonth', startMonth, 'endMonth', endMonth);

    // Cycle from September to June
    for (
        let currentMonth = new Date(startMonth.getTime()), previousLetterTotal, previousSheetName;
        isBefore(currentMonth, endMonth) || isSameMonth(currentMonth, endMonth);
        currentMonth = startOfMonth(addMonths(currentMonth, 1))
    ) {
        console.log('Processing', currentMonth)

        const sheetName = format(currentMonth, 'MMMM yyyy').toUpperCase();
        console.log('sheet name', sheetName);
        const currentSheet = workbook.addWorksheet(sheetName);

        const days = Array.from(new Array(30)) // parto da 31 giorni
            .map((e, i) => addDays(currentMonth, i)) // li aggiungo alla data iniziale del mese corrente
            .filter(day => { // poi filro fine settimana e giorni extra che non fanno parte del mese corrente
                if (day.getMonth() === 8 && day.getDate() < 14) { // september
                    return false;
                }
                if (day.getMonth() === 5 && day.getDate() > 10) { // june
                    return false;
                }
                return !isSunday(day) && isSameMonth(currentMonth, day);
            });

        // const daysHeader = days.map(day => ({header: format(day, 'dd/MM/yyyy'), width: 20}));
        const daysHeader = days.map(day => ({header: format(day, 'd-LLL').toUpperCase(), width: 15}));

        currentSheet.columns = [
            {header: 'DOCENTE', key: 'docente', width: 40},
            ...daysHeader,
            {width: 1},
            {header: 'TOTALE', key: 'totale', width: 12},
        ];
        currentSheet.addRow([undefined, ...days.map(day => format(day, 'EEEE').toUpperCase())])

        const lastDayLetter = currentSheet.getColumn(currentSheet.columns.length - 1).letter;

        teachersList.forEach((teacher, i) => {
            const rowIndex = 3 + (2 * i);
            let formula = `SUM(B${rowIndex}:${lastDayLetter}${rowIndex})`;

            if (previousLetterTotal && previousSheetName) {
                formula += ` + '${previousSheetName}'!${previousLetterTotal}${rowIndex}`;
            }

            currentSheet.addRow({docente: teacher.toUpperCase(), totale: {formula}});
            currentSheet.getRow(rowIndex).font = {size: 9, name: 'Calibri'}
            // currentSheet.getRow(rowIndex).fill = {
            //     type: 'pattern',
            //     pattern:'solid',
            //     fgColor:{argb:'CACACA'},
            // };
            currentSheet.getRow(rowIndex).alignment = {vertical: 'middle', horizontal: 'center'};
            // currentSheet.getRow(rowIndex).border = {
            //     top: {style:'thin'},
            //     left: {style:'thin'},
            //     bottom: {style:'thin'},
            //     right: {style:'thin'}
            // };
            currentSheet.getCell(`A${rowIndex}`).font = {size: 11, name: 'Calibri', bold: true,}
            currentSheet.getCell(`A${rowIndex}`).alignment = {vertical: 'middle', horizontal: 'left'}
            currentSheet.addRow();
            // currentSheet.getRow(rowIndex + 1).border = {
            //     top: {style:'thin'},
            //     left: {style:'thin'},
            //     bottom: {style:'thin'},
            //     right: {style:'thin'}
            // }
        });

        // Pin first row and column
        currentSheet.views = [
            {state: 'frozen', xSplit: 1, ySplit: 2}
        ];

        // Styles
        currentSheet.mergeCells('A1:A2');
        const alignment = {vertical: 'middle', horizontal: 'center'};
        currentSheet.getCell('A1:A2').alignment = alignment;
        const font = {
            size: 11, name: 'Calibri', bold: true,
        };
        currentSheet.getRow(1).font = font;
        currentSheet.getRow(2).font = font;
        currentSheet.getRow(1).alignment = alignment;
        currentSheet.getRow(2).alignment = alignment;
        currentSheet.getColumn(currentSheet.columns.length).font = font;
        currentSheet.getColumn(currentSheet.columns.length).alignment = alignment;
        const lastLetter = currentSheet.getColumn(currentSheet.columns.length).letter;
        currentSheet.mergeCells(`${lastLetter}1:${lastLetter}2`);

        // Borders and Color
        const len = currentSheet.columns.length;
        for (let i = 3; i < 3 + 2 * teachersList.length; ++i) {
            for (let j = 1; j <= len; ++j) {
                if (j !== len - 1) {
                    if (j !== len || currentSheet.getRow(i).getCell(1).value) {
                        currentSheet.getRow(i).getCell(j).border = {
                            top: {style:'thin'},
                            left: {style:'thin'},
                            bottom: {style:'thin'},
                            right: {style:'thin'}
                        }
                    }
                    if (currentSheet.getRow(i).getCell(1).value) {
                        currentSheet.getRow(i).getCell(j).fill = {
                            type: 'pattern',
                            pattern:'solid',
                            fgColor:{argb:'D2D2D2'},
                        };
                    }
                }


            }
        }

        previousLetterTotal = currentSheet.getColumn(currentSheet.columns.length).letter
        previousSheetName = sheetName;
    }

    // await workbook.xlsx.writeFile('output.xlsx');

    // write to a new buffer
    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    saveAs(blob, `Sostituzioni Permessi docenti ${year}_${year+1}.xlsx`);
}
