const fs = require('fs')
const XLSX = require('xlsx')
// local imports
const regexForStringSearch = require('./regexForStringSearch')
const createSheet = require('./createSheet')
const appendSheetsToWorkbook= require('./appendSheetsToWorkbook')

//////////////////////////////////////////////////
// Carrega o arquivo Excel
const planilha = XLSX.readFile('sheet.xlsx');

// Obtém a tabela desejada
const nomeDaTabela = planilha.SheetNames[0];
const tabela = planilha.Sheets[nomeDaTabela];
// Converte a tabela em JSON
const dadosJSON = XLSX.utils.sheet_to_json(tabela);

//==========================================================//
let sortedArr = []

const keywords = ['incorporadora','construtora','fundo','investidor','banco','industria','empresa',]

function basicParseSheet(){
  // combinar os nomeDaTabelas
  dadosJSON.forEach((contact) => {
      const combineName = [contact.name,contact.name2,contact.name3 ].join('')
      keywords.forEach(keyword => {
        if(combineName.match(regexForStringSearch(keyword))){
          const formatedContact = {
            nome:combineName,
            ramo: keyword,
            email:contact.Email,
            emailAlt:contact.emailAlt,
            phone1:contact.phone1,
            phone2:contact.phone2,
          }

          sortedArr.push(formatedContact)
        }
      })
  }
  )
}

// Filtra os dados
basicParseSheet()
const sheet = { nomeDaTabela: 'Filtrada', data: sortedArr }
createSheet(sheet.nomeDaTabela, sheet.data)

// Monta noco arquivo Excel
const wb = XLSX.utils.book_new();
appendSheetsToplanlha(wb, sheet);
XLSX.writeFile(wb, 'output.xlsx');


console.log(sortedArr)
