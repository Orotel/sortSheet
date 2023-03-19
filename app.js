const fs = require('fs');
const XLSX = require('xlsx');


// Carrega o arquivo Excel

const workbook = XLSX.readFile('sheet.xlsx');

// Obtém a aba desejada
const sheetName = workbook.SheetNames[0];
const workSheet = workbook.Sheets[sheetName];
// Converte as células em JSON
const dadosJSON = XLSX.utils.sheet_to_json(workSheet);

// Escreve o arquivo JSON
fs.writeFileSync('dados.json', JSON.stringify(dadosJSON));


//==========================================================//
  let incorporadora = []
  let construtora = []
  let fundo = []
  let investidor = []
  let banco = []
  let industria = []
  let empresa = [] 

const mycontact = {
  keywords: ['incorporadora','construtora','fundo','investidor','banco','industria','empresa',],
  

  regexForStringSearch(query) {
    query = query
    .replace(/a/gi, '[AÁÀÂÃ]')
      .replace(/e/gi, '[EÉÈÊ]')
      .replace(/i/gi, '[IÍÌÎ]')
      .replace(/o/gi, '[OÓÒÔÕ]')
      .replace(/u/gi, '[UÚÙÛ]')
      .replace(/c/gi, '[CÇ]')
      return new RegExp(query, 'gmi')
    },
    
    combinedName(){
      // combinar os sheetNames
      dadosJSON.forEach((contact) => {
        const combineName = [contact.name,contact.name2,contact.name3 ].join('')
        this.keywords.forEach(keyword => {
        if(combineName.match(this.regexForStringSearch(keyword))){
          const formatedContact = {
            nome:combineName,
            email:contact.Email,
            emailAlt:contact.emailAlt,
            phone1:contact.phone1,
            phone2:contact.phone2,
          }

          switch(keyword) {
            case "incorporadora":
              incorporadora.push(formatedContact);
              break;
            case "construtora":
              construtora.push(formatedContact);
              break;
            case "fundo":
              fundo.push(formatedContact);
              break;
            case "investidor":
              investidor.push(formatedContact);
              break;
            case "banco":
              banco.push(formatedContact);
              break;
            case "industria":
              industria.push(formatedContact);
              break;
            case "empresa":
              empresa.push(formatedContact);
              break;
            default:
              break
          }        
        }
      })
    }
    )
 },
}
// Adicione seus objetos de contato aos arrays aqui

function createSheet(sheetName, data) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return wb;
}

function appendSheetsToWorkbook(workbook, sheets) {
  sheets.forEach(({ sheetName, data }) => {
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  });
  return workbook;
}

const sheets = [
  { sheetName: 'Incorporadora', data: incorporadora },
  { sheetName: 'Construtora', data: construtora },
  { sheetName: 'Fundo', data: fundo },
  { sheetName: 'Investidor', data: investidor },
  { sheetName: 'Banco', data: banco },
  { sheetName: 'Industria', data: industria },
  { sheetName: 'Empresa', data: empresa },
];

mycontact.combinedName()
const wb = XLSX.utils.book_new();
appendSheetsToWorkbook(wb, sheets);

// Substitua 'output.xlsx' pelo nome desejado do arquivo
XLSX.writeFile(wb, 'output.xlsx');

createSheet(`batata`,fundo)
console.log(incorporadora)
