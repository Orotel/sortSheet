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
          switch(keyword) {
            case "incorporadora":
              incorporadora.push(contact);
              break;
            case "construtora":
              construtora.push(contact);
              break;
            case "fundo":
              fundo.push(contact);
              break;
            case "investidor":
              investidor.push(contact);
              break;
            case "banco":
              banco.push(contact);
              break;
            case "industria":
              industria.push(contact);
              break;
            case "empresa":
              empresa.push(contact);
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



mycontact.combinedName()
// console.log(construtora)