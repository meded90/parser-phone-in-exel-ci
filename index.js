const { findPhoneNumbersInText } = require('libphonenumber-js')
const Excel = require('exceljs');


// ------------------------------------------
// constant for use
//
const PATH_FILE = 'Только_ячейка_с_телефонами_Дадата.xlsx'
const PAGE_NAME = 'Лист1'
const TARGET_CELL = 2
const RESULT_CELL = 11

function remover(str){
  return str
    ?.replace('+ +', '+')
    .replace('+8', '8')
    .replace(/\+$/, '')
    .replace(/^\+/, '')
    .replace('+ 8', '8')
    .replace(' ,89', ', 89')
    .replace(')8', ') 8')
    .replace('++7', '8')
    .replace('+79', ' 89')
    .replace(',9', ', 89')
    // цены
    .replace('3500,', '')
    .replace(' 3500 ', '')
    .replace('2990-', '')
    .replace(' 2990 ', '')
    .replace('2990,', '')
    .replace(' 1600', '')
    .replace(' 1100', '')
    .replace('1990,', '')
    .replace(' 2500,', '')
    .replace('1990 ,', '')
    .replace('1990+', '')
    .replace('+1990', '')
    .replace('1400,', '')
    .replace('1100,', '')
    .replace('1600,', '')
    .replace('1000,', '')
    .replace(/1000$/, '')
    .replace('800,', '')
    .replace('брови 3990', '')
    .replace('брови 4500', '')
    .replace('3500 брови', '')
    .replace('брови 3500', '')
    .replace('брови 1990', '')
    .replace('брови 2990', '')
    .replace('брови 1600', '')
    .replace('бровки 1600', '')
    .replace('брови 1000', '')
    .replace('2990 брови', '')
    .replace('1990 брови', '')
    .replace('губы 1990', '')
    .replace(' 2500 губы', '')
    .replace('губы 2500', '')
    .replace('1000 брови', '')
    .replace('брови 2500', '')
    .replace('брови 3990', '')
    .replace('3990 брови', '')
    .replace('корр.3990', '')
    .replace('1000 корр', '')
    .replace('2500 корр', '')
    .replace('КОРРЕКЦИЯ 1990', '')
    .replace('коррекция 2990', '')
    .replace('корр 2990', '')
    .replace('корр.3591', '')
    .replace('корр 1990', '')
    .replace('хна 1100', '')
    .replace('модели 1000', '')
    .replace('веки 3990', '')
    .replace('2990 веки', '')
    .replace('2990 веки', '')
    .replace('1600 резерв', '')
    .replace('повторная 2500', '')
    .replace('соседка 1990', '')
    .replace('по 2500', '')
    .replace('наталья 29908', '')
    .replace(' К 4990 ', '')
    .replace('10%', '')
    .replace('-10%', '')
    .replace('1990-30%=1400', '')
    .replace('2022+2', '')
    .replace('6066+6', '')
    .replace('2034+3', '')
    .replace('нас 3990', '')
    // даты
    .replace('9.06', '')
    .replace('17.07,', '')
    .replace('30.06,', '')
    .replace('14 .07', '')
    .replace('29.05,', '')
    .replace('26.08,', '')
    .replace('(09.01 ', '')
    .replace('20.03,', '')
    .replace('23.06,', '')
    .replace(' 27.05,', '')
    .replace('25.08,', '')
    .replace('27.05, ', '')
    .replace('27.05 ,', '')
    .replace('10 и 12.30', '')

}

// ------------------------------------------
// init
//
const wb = new Excel.Workbook();
const path = require('path');
const filePath = path.resolve(__dirname, PATH_FILE);
process.exitCode = 1

// ------------------------------------------
// parsing
//
async function start() {

  await wb.xlsx.readFile(filePath)

  const worksheet = wb.getWorksheet(PAGE_NAME);
  worksheet.eachRow(function (row, rowNumber) {

    let targetContent = row.getCell(TARGET_CELL).value;
    // дополнительные парфиры
    targetContent = remover(targetContent)


    const phoneNumber = findPhoneNumbersInText(targetContent, 'RU')

    if (!phoneNumber.length) {
      console.log(`%c not find row:${ rowNumber }  ${ targetContent } `, ` color: #f82121 `);
      return
    }

    phoneNumber.forEach((phone, index) => {
      let value = phone.number.format('E.164');
      row.getCell(RESULT_CELL + index).value = value
      // console.log(`row:`, rowNumber, '->', value);
    })

  })

  await wb.xlsx.writeFile(filePath);
}

start();



