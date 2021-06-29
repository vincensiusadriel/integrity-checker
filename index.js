const xlsx = require('xlsx')
const fs = require('fs')

let workbook = xlsx.readFile('Book1.xlsx');
let sheet_name_list = workbook.SheetNames;
let xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

let n = xlData.length

let final = ''
let prev = ''
for (let i = 0; i < n; i++) {
    let obj = xlData[i]

    let tableName = obj.tableName
    let columnName = obj.columnName
    let joinTableName = obj.joinTableName
    let joinColumnName = obj.joinColumnName
    let refMasterType = obj.refMasterType

    let res = ''

    if (prev != tableName) {
        res = `--==================   ${tableName}   ============================\n\n`
        prev = tableName
    }

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
                SELECT COUNT(${tableName}_ID) AS COUNTID, ${columnName} 
                FROM ${tableName} WITH ( NOLOCK )
                GROUP BY ${columnName}
                HAVING COUNT(${tableName}_ID) > 1
                ) X
                `
            break;

        case 'RM':
            res += `
                SELECT '${tableName}.${columnName}' AS [TABLE] ,
                ${tableName}.${columnName} AS VALUE
                FROM dbo.${tableName} WITH(NOLOCK)
                LEFT JOIN dbo.REF_MASTER WITH(NOLOCK) ON REF_MASTER.MASTER_CODE = ${columnName} AND REF_MASTER.REF_MASTER_TYPE_CODE = '${refMasterType}'
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
                WHERE ${joinTableName}_ID IS NULL
                GROUP BY ${tableName}.${columnName}
                `
            break;


    }

    final += res + '\n\n'

}


fs.writeFileSync('result.txt', final)