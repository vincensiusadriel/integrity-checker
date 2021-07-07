const xlsx = require('xlsx')
const prompt = require('prompt-sync')()
const fs = require('fs')

let filename = prompt('Write filename without the extension (.xlsx) (default : Book1) : ')
console.log(filename)
if (filename == '') {
    filename = 'Book1'
}

let workbook = xlsx.readFile(filename + '.xlsx');
let sheet_name_list = workbook.SheetNames;
let xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

let n = xlData.length


let gigaType = prompt('Result type (c: by case (default), t: by table with no insert) : ')
let final = ''
let prev = ''


const query = (obj, res) => {
    let tableName = obj.tableName
    let columnName = obj.columnName
    let joinTableName = obj.joinTableName
    let joinColumnName = obj.joinColumnName
    let grpCode = obj.grpCode

    switch (obj.type) {
        case 'IS':
            res += `
                Select    '${tableName}.${columnName}' As [TABLE] ,
                ${tableName}.${columnName} As VALUE
                From      ${tableName} With ( NoLock )
                Where     ${columnName} Not In ( 1, 0 )
                Or ${columnName} Is Null
                `
            break;

        case 'DUP':
            res += `
                SELECT    '${tableName}.${columnName}' AS [TABLE] ,
                X.${columnName} AS VALUE
                FROM      (
                SELECT COUNT(1) AS COUNTID, ${columnName} 
                FROM ${tableName} WITH ( NOLOCK )
                GROUP BY ${columnName}
                HAVING COUNT(1) > 1
                ) X
                `
            break;

        case 'RM':
            res += `
                SELECT '${tableName}.${columnName}' AS [TABLE] ,
                ${tableName}.${columnName} AS VALUE
                FROM dbo.${tableName} WITH(NOLOCK)
                LEFT JOIN dbo.REF_MASTER WITH(NOLOCK) ON REF_MASTER.MASTER_CODE = ${columnName} AND REF_MASTER.REF_MASTER_TYPE_CODE = '${grpCode}'
                WHERE REF_MASTER_ID IS NULL
                GROUP BY ${columnName}
                `
            break;

        case 'JOIN':
            res += `
                SELECT '${tableName}.${columnName}' AS [TABLE] ,
                ${tableName}.${columnName} AS VALUE
                FROM dbo.${tableName} WITH(NOLOCK)
                LEFT JOIN dbo.${joinTableName} WITH(NOLOCK) ON ${joinTableName}.${joinColumnName} = ${tableName}.${columnName}
                WHERE ${joinTableName}.${joinColumnName} IS NULL
                GROUP BY ${tableName}.${columnName}
                `
            break;


    }
    return res
}
fs.writeFileSync('result.txt', final)
if (gigaType == 't') {
    for (let i = 0; i < n; i++) {
        let obj = xlData[i]



        let res = ''

        if (prev != obj.tableName) {
            res = `--==================   ${obj.tableName}   ============================\n\n`
            prev = obj.tableName
        }

        res = query(obj, res)

        fs.appendFileSync('result.txt', res + '\n\n')
    }
} else {
    for (let i = 0; i < n; i++) {
        let obj = xlData[i]

        let res = ''

        if (prev != obj.type) {
            let typeString = ''

            switch (obj.type) {
                case 'IS':
                    typeString = 'IS_ DIISI 0 ATAU 1, BUKAN Y ATAU N, TIDAK BOLEH NULL'
                    break;

                case 'DUP':
                    typeString = 'DATA TIDAK BOLEH DUPLIKAT'
                    break;

                case 'RM':
                    typeString = 'DATA MR_ BELUM TERDAFTAR DI REF_MASTER'
                    break;

                case 'JOIN':
                    typeString = 'DATA TIDAK DITEMUKAN DI TABEL JOIN'
                    break;


            }


            let begin = `
        --====================================== ${typeString}
    
        Insert  Into INTEGRITY_CHECK
            ( ISSUE ,
              [TABLE] ,
              VALUE
            )
            Select  '${typeString}' ,
                    [TABLE] ,
                    VALUE
            From    (
        `
            res += begin
            prev = obj.type
        }

        res = query(obj, res)

        if (xlData[i + 1] == undefined || obj.type != xlData[i + 1].type) {
            res += `\n        ) AS X`
        } else {
            res += '\n                UNION\n'
        }
        fs.appendFileSync('result.txt', res)
    }
}


console.log('Success !')