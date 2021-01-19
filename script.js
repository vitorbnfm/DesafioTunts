const { GoogleSpreadsheet } = require('google-spreadsheet');
const credentials = require('./credentials.json')


// Initialize the sheet - doc ID is the long id in the sheets URL
const doc = new GoogleSpreadsheet('1OB0mnUo0TkHNEj9_0VH53jGIS9EGwDWMYqHyxKL3iYc');


async function accessSpreadsheet() {

    // Initialize Auth
    await doc.useServiceAccountAuth(credentials, err => {
        console.log(err);
    });

    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    await sheet.loadCells();
    const row = await sheet.getRows();
    console.log(row.length)

    for (let i = 0; i <= row.length; i++) {

        const id = sheet.getCellByA1(`A${4 + i}`);
        console.log(id.value)
        if (id.value == null) {
            break;
        }


        const studentName = sheet.getCellByA1(`B${4 + i}`);
        console.log(`Aluno: ${studentName.value}`);


        const attendance = sheet.getCellByA1(`C${4 + i}`);
        console.log(`Faltas: ${attendance.value}`);


        const firstExam = sheet.getCellByA1(`D${4 + i}`);
        const secondExam = sheet.getCellByA1(`E${4 + i}`);
        const thirdExam = sheet.getCellByA1(`F${4 + i}`);


        const average = ((firstExam.value + secondExam.value + thirdExam.value) / 3).toFixed(0);
        console.log(`Média: ${average}`);


        const situation = sheet.getCellByA1(`G${4 + i}`);
        const requiredGrade = sheet.getCellByA1(`H${4 + i}`);


        if (average < 50) {
            situation.value = 'Reprovado por nota';
            requiredGrade.value = 0;
            console.log('Reprovado por nota');
        }
        else if (average >= 50 && average < 70) {
            situation.value = 'Exame Final';
            const grade = 100 - average;
            requiredGrade.value = grade;
            console.log('Necessario exame final');
        }
        else if (average >= 70) {
            situation.value = 'Aprovado';
            requiredGrade.value = 0;
            console.log('Aprovado');
        }
        if (attendance.value > 15) {
            situation.value = 'Reprovado por Falta';
            requiredGrade.value = 0;
            console.log('Reprovado por Falta');
        }
        console.log('--------------------------------');
        await sheet.saveUpdatedCells();
    }
}

accessSpreadsheet()