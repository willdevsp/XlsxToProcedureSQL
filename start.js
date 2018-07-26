const XLSX = require('xlsx');
const fs = require('fs');
var readlineSync = require('readline-sync');

var xlData = null;
var filesInFolder = {};


var arquivos = fs.readdirSync("./arquivos", function (err, files) {

});


var listFile = () => {
    console.log('Arquivos XLSX');
    arquivos.map((nome, posicao) => {
        filesInFolder[posicao] = nome;
        console.log(`${posicao} - ${nome}`);
    });
    return '';
}



var store_id = readlineSync.question(`Qual Store_ID?`);
store_id = store_id.toUpperCase();
var proc = readlineSync.question('Qual procedure? ');
var position = readlineSync.question(`Escolha um arquivo a ser usado?\n${listFile()} `);
var xlsx_file = './arquivos/' + filesInFolder[position];
var nome_aba = readlineSync.question(`Qual o nome da aba do arquivo ${xlsx_file} ?`);
var sql_file = readlineSync.question('Qual o nome do arquivo SQL a ser gerado? ');
console.log(`Store_id: ${store_id}\nProcedure:${proc}\nArquivoXLS: ${xlsx_file}\nAba:${nome_aba}\nArquivoSQL: ${sql_file} ?`);
var confirmar = readlineSync.question('Deseja confirmar as alterações [S]im, [N]ão ?');

var createSQL = (proc, store_id, xlsx_file, sql_file, nome_aba) => {
    var workbook = XLSX.readFile(`${xlsx_file}`);
    xlData = XLSX.utils.sheet_to_json(workbook.Sheets[nome_aba]);
    var dateFile = new Date(Date.now());
    var dateFileFormatted = dateFile.toISOString();
    sql_file = `./sql_gerados/${sql_file}${dateFileFormatted}_${store_id}.sql`;

    if (fs.existsSync(sql_file)) {
        fs.unlink(sql_file, function (err) {
            if (err) throw err;
            console.log('File deleted!');
        });
    }

    const addToSQL = query => {
        fs.appendFileSync(sql_file, query, function (err) {
            if (err) throw err;
        });
    }    
    var quantity = xlData.length-1;

    xlData.forEach((data, cont) => {
        var x = Object.values(data);
        values = x.map(value => {
            return `'${value}'`;
        });
        if (cont == 0) {
            addToSQL(`BEGIN\nBEGIN ${proc} (${values}, '${store_id}'); END;\n`);
        } else if (cont == quantity) {
            addToSQL(`BEGIN ${proc} (${values}, '${store_id}'); END;\nCOMMIT; \nEND;`);
        } else {
            addToSQL(`BEGIN ${proc} (${values}, '${store_id}'); END;\n`);
        }
    })
    console.log(`${sql_file} gerado com sucesso`);
}


if (confirmar == 's' || confirmar == 'S') {
    console.log('Gerando arquivo SQL, aguarde...');
    createSQL(proc, store_id, xlsx_file, sql_file, nome_aba);
} else {
    console.log('Abortado');
}






