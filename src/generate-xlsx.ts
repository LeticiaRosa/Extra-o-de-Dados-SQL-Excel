import OracleDB from 'oracledb'
import Database from './Database'
import ExcelJs from 'exceljs'
import path from 'node:path'
import fs from 'node:fs'
async function main() {
  const conection = await Database.getInstance().connect('usrgepe', 'usrgepe01')
  await regra()
  await Database.getInstance().disconnect()
}

main()
const query = `
select case when RHPESS_CONTRATO.codigo_empresa = '0001' then 'ADMINSTRAÇÃO DIRETA' else  RHORGA_EMPRESA.razao_social end  orgao,
    case when trim(unidade1.texto_associado) is null then RHORGA_EMPRESA.razao_social else trim(unidade1.texto_associado) end as Nivel_Hierarquico,
	   UPPER(TRIM(RHPESS_CONTRATO.nome)) Nome,
     case when NVL(RHPESS_CONTRATO.cod_cargo_efetivo,'000000000000000') <> '000000000000000' then UPPER(TRIM(cargo_a.descricao))
			  else ''
		end Cargo,
	   NVL(sum(movimento.remuneracao),'') total_remuneracao
from RHPESS_CONTRATO,
	 rhplcs_cargo cargo_a,
	 rhplcs_funcao,
   rhparm_sit_func,
    RHORGA_EMPRESA,
	 (
		select RHMOVI_MOVIMENTO.codigo_empresa codigo_empresa,
			   RHMOVI_MOVIMENTO.codigo_contrato codigo_contrato,
			   RHMOVI_MOVIMENTO.tipo_contrato tipo_contrato,
				RHMOVI_MOVIMENTO.ano_mes_referencia ano_mes_referencia,
				RHMOVI_MOVIMENTO.valor_verba as remuneracao

		from RHMOVI_MOVIMENTO
		where RHMOVI_MOVIMENTO.tipo_movimento = 'ME' and
			  RHMOVI_MOVIMENTO.fase = '0' and
			  RHMOVI_MOVIMENTO.modo_operacao = 'R'and
			  RHMOVI_MOVIMENTO.valor_verba > '0' and
			  RHMOVI_MOVIMENTO.ano_mes_referencia = TO_DATE(:data_execute, 'dd/mm/yyyy') and
        ( 
        ( RHMOVI_MOVIMENTO.codigo_verba in ('3400') and codigo_empresa in('0001'))  or 
        ( RHMOVI_MOVIMENTO.codigo_verba in ('1B2M','1B2O','4BB1') and codigo_empresa = '0009') or 
        ( RHMOVI_MOVIMENTO.codigo_verba in ('1135','3400','1F37') and codigo_empresa = '0013')or
        ( RHMOVI_MOVIMENTO.codigo_verba in ('1135','3400') and codigo_empresa = '0014')  or 
        ( RHMOVI_MOVIMENTO.codigo_verba in ('4P2F','1P2N') and codigo_empresa = '0002' ) or
        ( RHMOVI_MOVIMENTO.codigo_verba in ('4S76') and codigo_empresa = '0007' ) or 
        ( RHMOVI_MOVIMENTO.codigo_verba in ('10H6','1135','50A8') and codigo_empresa = '0003' )or 
        ( RHMOVI_MOVIMENTO.codigo_verba in ('4U3K') and codigo_empresa = '0010' and rhmovi_movimento.ctrl_demo = 'N' )or 
        ( RHMOVI_MOVIMENTO.codigo_verba in ('4001') and codigo_empresa = '0005')          
        )
	 ) movimento,

	 (
		select rhorga_unidade.codigo_empresa codigo_empresa,
				 rhorga_unidade.cod_unidade1 cod_unidade1,
				 UPPER(TRIM(NVL(TRIM(rhorga_unidade.texto_associado),rhorga_unidade.descricao))) texto_associado
		from rhorga_unidade
		where rhorga_unidade.cod_unidade2 = '000000' and
				rhorga_unidade.cod_unidade3 = '000000' and
				rhorga_unidade.cod_unidade4 = '000000' and
				rhorga_unidade.cod_unidade5 = '000000' and
				rhorga_unidade.cod_unidade6 = '000000'
	 ) unidade1

where 
	  RHPESS_CONTRATO.codigo = movimento.codigo_contrato and
	  RHPESS_CONTRATO.codigo_empresa = movimento.codigo_empresa and
	  RHPESS_CONTRATO.tipo_contrato = movimento.tipo_contrato and
	  RHPESS_CONTRATO.cod_cargo_efetivo = cargo_a.codigo and
	  RHPESS_CONTRATO.codigo_empresa = cargo_a.codigo_empresa and
    RHPESS_CONTRATO.situacao_funcional = rhparm_sit_func.codigo and

	  RHPESS_CONTRATO.codigo_funcao = rhplcs_funcao.codigo(+) and
	  RHPESS_CONTRATO.codigo_empresa = rhplcs_funcao.codigo_empresa(+) and


	  RHPESS_CONTRATO.codigo_empresa = unidade1.codigo_empresa  and
	  RHPESS_CONTRATO.cod_unidade1 = unidade1.cod_unidade1 and

	  rhorga_empresa.codigo = RHPESS_CONTRATO.codigo_empresa and
    RHPESS_CONTRATO.vinculo = '0009' and
	  RHPESS_CONTRATO.codigo_empresa in ('0001','0002','0003','0005','0007','0009','0010','0013','0014') and
	  rhparm_sit_func.controle_folha in ('N','L','M') AND
	  RHPESS_CONTRATO.ano_mes_referencia = (select max (a.ano_mes_referencia)
											from RHPESS_CONTRATO a
											where a.codigo = RHPESS_CONTRATO.codigo and
												  a.codigo_empresa = RHPESS_CONTRATO.codigo_empresa and
												  a.tipo_contrato = RHPESS_CONTRATO.tipo_contrato and
												  a.ano_mes_referencia <= TO_DATE(:data_execute, 'dd/mm/yyyy')) 

group by 
		RHPESS_CONTRATO.codigo_empresa,
    RHPESS_CONTRATO.tipo_contrato,
		RHORGA_EMPRESA.razao_social,
	   UPPER(TRIM(RHPESS_CONTRATO.nome)) ,
    case when trim(unidade1.texto_associado) is null then RHORGA_EMPRESA.razao_social else trim(unidade1.texto_associado) end ,
		case when NVL(RHPESS_CONTRATO.cod_cargo_efetivo,'000000000000000') <> '000000000000000' then UPPER(TRIM(cargo_a.descricao))
			  else ''
		end`

const directoryPath = './arquivos'
async function regra() {
  const years = [2021, 2022, 2023, 2024]
  const months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
  for (const year of years) {
    for (const month of months) {
      const dataFormatada: string = `01/${month.toString().padStart(2, '0')}/${year}`
      const result = await Database.getInstance().executeQuery(query, {
        data_execute: {
          dir: OracleDB.BIND_IN,
          type: OracleDB.STRING,
          val: dataFormatada,
        },
      })

      const mesAnoFile = `${month.toString().padStart(2, '0')}-${year}`
      await generateFile(
        result,
        mesAnoFile,
        year.toString()
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
) {
  const columns = [
    {
      header: 'ÓRGÃO',
      key: 'ORGAO',
    },
    {
      header: 'NIVEL HIERÁRQUICO',
      key: 'NIVEL_HIERARQUICO',
    },
    {
      header: 'NOME',
      key: 'NOME',
    },
    {
      header: 'CARGO',
      key: 'CARGO',
    },
    {
      header: 'TOTAL REMUNERAÇÃO',
      key: 'TOTAL_REMUNERACAO',
    },
  ]
  const workbook = new ExcelJs.Workbook()
  workbook.creator = 'Dados Abertos'
  workbook.created = new Date()
  const file = workbook.addWorksheet('DADOS')
  file.columns = columns
  file.addRows(data)

  if (ano) {
    const buffer = (await workbook.xlsx.writeBuffer()) as Buffer
    const yearDirectory = path.join(
      directoryPath,
      ano,
    )
    if (!fs.existsSync(yearDirectory)) {
      fs.mkdirSync(yearDirectory, { recursive: true })
    }

    fs.writeFileSync(
      path.join(yearDirectory, `Relatorio_Nominal_Estagiarios_Remuneracao_${mesAnoFile}.xlsx`),
      buffer,
    )
  }
}
