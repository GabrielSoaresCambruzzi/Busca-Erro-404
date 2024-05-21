const axios = require('axios')
const XLSX = require('xlsx')
const fs = require('fs')

function lerLinks(caminhoArquivo) {
    const workbook = XLSX.readFile(caminhoArquivo)
    const sheetName = workbook.SheetNames[0]; //Lê a primeira aba
    const sheet = workbook.Sheets[sheetName]
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1})
    const links = data.flat().filter(cell => typeof cell === 'string' && cell.startsWith('http')) // filtrar os links da planilha
    return links
}

// Verificar se um link retorna erro 404
async function verificarErro(url) {
    try {
        const response = await axios.get(url)
        return response.status === 404    
        } catch (error) {
            if (error.response && error.response.status === 404) {
                return true
            }
            return false
        }
    }

// Salvar links em uma nova planilha
function salvarLink(caminhoArquivo, links) {
    const workbook = XLSX.utils.book_new()
    const sheet = XLSX.utils.aoa_to_sheet([["Links com erro 404"], ...links.map(link => [link])])
    XLSX.utils.book_append_sheet(workbook, sheet, 'Erro 404')
    XLSX.writeFile(workbook, caminhoArquivo)
}

async function processarLinks(caminhoEntrada, caminhoSaida) {
    const links = lerLinks(caminhoEntrada)
    const linkscomErro = []

    for (const link of links) {
        const temErro404 = await verificarErro(link)
        if (temErro404) {
            linkscomErro.push(link)
        }
    }
    salvarLink(caminhoSaida, linkscomErro)
    console.log(`Processamento concluido. ${linkscomErro.length} links com erro 404 foram salvos em ${caminhoSaida}.`)
}

// Caminhos dos arquivos
const caminhoEntrada = 'teste.xlsx'; // Substitua pelo caminho da sua planilha de entrada
const caminhoSaida = 'links_com_erro_404.xlsx'; // Substitua pelo caminho da sua planilha de saída

// Execute o processamento
processarLinks(caminhoEntrada, caminhoSaida);






/*async function verificarErro(url) {
    try{
        const response = await axios.get(url)
        if (response.status === 404) {
            console.log(`A URL ${url} retornou um erro 404`)
        } else{
            console.log(`A URL ${url} retornou o status code ${response.status}`)
        }
    } catch (error) {
        if (error.response) {
            console.log(`A URL ${url} retornou o status code: ${error.response.status}`)
        } else if (error.resquest) {
            console.log(`Nenhuma resposta recebida para URL ${url}`)
        } else {
            console.log(`Ocorreu um erro ao tentar acessar a URL ${url}: ${error.message}`)
        }
    }
}

const url = "https://static.weg.net/medias/downloadcenter/he4/hfd/WEG-system-automation-solutions-50052180-en.pdf"
verificarErro(url)
*/