import OracleDB from 'oracledb'
import Database from './Database'
import ExcelJs from 'exceljs'
import path from 'node:path'
import fs from 'node:fs'
async function main() {
  await Database.getInstance().connect('usrgepe', 'usrgepe01')
  await runQuery()
  await Database.getInstance().disconnect()
}

main()
const query = `SELECT   
 emp.razao_social as "ENTIDADE/ÓRGÃO",
   c.nome,  
   CASE  
     WHEN efet.descricao IS NULL or substr(efet.codigo,12,4) in ( '0000' ) 
     THEN ''  
     ELSE efet.descricao  
   END AS cargo_efetivo,  
   CASE  
     WHEN comiss.descricao IS NULL or substr(comiss.codigo,12,4) in ( '0000' )
     THEN ''  
     ELSE comiss.descricao  
   END AS cargo_comissionado,
    CASE  
     WHEN func.DESCRICAO IS NULL or substr(func.codigo,12,4)  in ( '0000' )
     THEN ''  
     ELSE func.DESCRICAO  
   END AS função,
 SUBSTR( c.COD_CUSTO_GERENC1,5,2)   
   ||'.'   
   || UPPER(TRIM(SUBSTR( c.COD_CUSTO_GERENC2,5,2)) )   
   ||'.'   
   || UPPER(TRIM(SUBSTR( c.COD_CUSTO_GERENC3,5,2)) )   
   ||'.'   
   || UPPER(TRIM(SUBSTR( c.COD_CUSTO_GERENC4,5,2)) )   
   ||'.'   
   || UPPER(TRIM(SUBSTR( c.COD_CUSTO_GERENC5,5,2)) )   
   ||'.'   
   || UPPER(TRIM(SUBSTR( c.COD_CUSTO_GERENC6,5,2)) ) ||' - '  ||
    g.TEXTO_ASSOCIADO AS "UNIDADE DE LOTAÇAO",
 
    case when jor.C_LIVRE_DESCR02 = 'DIARIA' then to_number(nvl(trunc(JOR.c_livre_selec01),0))*5 else to_number(nvl(trunc(JOR.c_livre_selec01),0))  end as "CARGA HORÁRIA SEMANAL EM MINUTOS" ,
   C.DATA_ADMISSAO,
   CASE WHEN sf.controle_folha <> 'S' THEN C.DATA_RESCISAO END AS DATA_RESCISAO,
   CASE WHEN  sf.controle_folha = 'S' THEN C.DATA_RESCISAO END AS DATA_INATIVACAO
 FROM rhpess_contrato c
 inner join RHORGA_EMPRESA emp
 on c.codigo_empresa = emp.codigo
 inner join RHORGA_CUSTO_GEREN g on
  c.COD_CUSTO_GERENC1   = g.COD_CGERENC1   
 AND c.COD_CUSTO_GERENC2   = g.COD_CGERENC2   
 AND c.COD_CUSTO_GERENC3   = g.COD_CGERENC3   
 AND c.COD_CUSTO_GERENC4   = g.COD_CGERENC4   
 AND c.COD_CUSTO_GERENC5   = g.COD_CGERENC5   
 AND c.COD_CUSTO_GERENC6   = g.COD_CGERENC6   
 AND c.codigo_empresa      = g.codigo_empresa   
 inner join rhplcs_cargo efet  on  
  efet.codigo         = c.cod_cargo_efetivo  
 AND efet.codigo_empresa = c.codigo_empresa  
 left join rhplcs_cargo comiss   on 
  comiss.codigo          = c.cod_cargo_comiss  
 AND comiss.codigo_empresa  = c.codigo_empresa   
 inner join rhpont_escala escala  on 
 escala.codigo         = c.codigo_escala  
 AND escala.codigo_empresa = c.codigo_empresa 
 INNER JOIN RHPONT_TP_JORNADA jor
 ON jor.codigo              = escala.tipo_jornada
 left join RHPLCS_FUNCAO func  on 
  c.CODIGO_FUNCAO  = func.CODIGO 
 AND c.codigo_empresa = func.CODIGO_EMPRESA  
 inner join rhparm_sit_func sf
 on sf.codigo             = c.situacao_funcional   
 WHERE c.ANO_MES_REFERENCIA =  
   (SELECT MAX (a.ano_mes_referencia)  
   FROM rhpess_contrato a  
   WHERE a.codigo       = c.codigo  
   AND a.codigo_empresa = c.codigo_empresa  
   AND a.tipo_contrato  = c.tipo_contrato  
   AND a.ano_mes_referencia <= TO_DATE(:data_execute, 'dd/mm/yyyy') )   
 and ( c.data_rescisao is null OR TO_DATE(TO_CHAR(C.data_rescisao, 'mm/yyyy'), 'mm/yyyy') = TO_DATE(:mes_ano, 'mm/yyyy') ) 
 and (( c.codigo_empresa in  ('0001')  AND 'S'= :is_adm_direta ) OR 
   ( c.codigo_empresa in ('0002','0003','0005','0007','0009','0010','0013','0014')  AND 'N'= :is_adm_direta ) )
 ORDER BY 1,2`
async function runQuery() {
  const admDireta = ['0001']
  const admIndireta = [
    '0002',
    '0003',
    '0005',
    '0007',
    '0009',
    '0010',
    '0013',
    '0014',
  ]
  for (let i = 0; i < 2; i++) {
    if (i === 0) {
      await regra(admDireta)
    } else {
      await regra(admIndireta)
    }
  }
}
const directoryPath = './arquivos'

async function regra(empresa: string[]) {
  const years = [2021, 2022, 2023, 2024]
  const months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
  for (const year of years) {
    for (const month of months) {
      const dataFormatada: string = `01/${month.toString().padStart(2, '0')}/${year}`
      console.log(dataFormatada)
      const result = await Database.getInstance().executeQuery(query, {
        data_execute: {
          dir: OracleDB.BIND_IN,
          type: OracleDB.STRING,
          val: dataFormatada,
        },
        mes_ano: {
          dir: OracleDB.BIND_IN,
          type: OracleDB.STRING,
          val: dataFormatada.split('/')[1] + '/' + dataFormatada.split('/')[2],
        },
        is_adm_direta: {
          dir: OracleDB.BIND_IN,
          type: OracleDB.STRING,
          val: empresa[0] === '0001' ? 'S' : 'N',
        },
      })

      const mesAnoFile = `${month.toString().padStart(2, '0')}-${year}`
      await generateFile(
        result,
        mesAnoFile,
        year.toString(),
        empresa[0] === '0001',
      )

      if (year === 2024 && month > new Date().getMonth() - 1) {
        break
      }
    }
  }
}

async function generateFile(
  data: unknown[],
  mesAnoFile: string,
  ano: string,
  isDireta: boolean,
) {
  const columns = [
    {
      header: 'ENTIDADE/ÓRGÃO',
      key: 'ENTIDADE/ÓRGÃO',
    },
    {
      header: 'NOME',
      key: 'NOME',
    },
    {
      header: 'CARGO_EFETIVO',
      key: 'CARGO_EFETIVO',
    },
    {
      header: 'CARGO_COMISSIONADO',
      key: 'CARGO_COMISSIONADO',
    },
    {
      header: 'FUNÇÃO',
      key: 'FUNÇÃO',
    },
    {
      header: 'UNIDADE DE LOTAÇÃO',
      key: 'UNIDADE DE LOTAÇÃO',
    },
    {
      header: 'CARGA HORÁRIA SEMANAL EM MINUTOS',
      key: 'CARGA HORÁRIA SEMANAL EM MINUTOS',
    },
    {
      header: 'DATA_ADMISSAO',
      key: 'DATA_ADMISSAO',
    },
    {
      header: 'DATA_RESCISAO',
      key: 'DATA_RESCISAO',
    },
    {
      header: 'DATA_INATIVACAO',
      key: 'DATA_INATIVACAO',
    },
  ]
  const workbook = new ExcelJs.Workbook()
  workbook.creator = 'Sistema de Transferência'
  workbook.created = new Date()
  const file = workbook.addWorksheet('Dados')
  file.columns = columns
  file.addRows(data)

  if (ano) {
    const buffer = (await workbook.xlsx.writeBuffer()) as Buffer

    const yearDirectory = path.join(
      directoryPath,
      ano,
      isDireta ? 'DIRETA' : 'INDIRETA',
    )
    if (!fs.existsSync(yearDirectory)) {
      fs.mkdirSync(yearDirectory, { recursive: true })
    }

    fs.writeFileSync(
      path.join(yearDirectory, `Dados_${mesAnoFile}.xlsx`),
      buffer,
    )
  }
}
