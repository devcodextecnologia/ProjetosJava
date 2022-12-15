package br.com.bulla.LeitorPlanilha;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Arquivo {

	
	public void criarListaDados(List<Dados> dados, String nomeArquivo) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Dados");

		int linha = 0;

			
		for (Dados dado : dados) {

			Row row = sheet.createRow(linha++);

			int celula = 0;
			Cell OPERACAO = row.createCell(celula++);
			OPERACAO.setCellValue(dado.getOPERACAO());

			Cell CONVENIO = row.createCell(celula++);
			CONVENIO.setCellValue(dado.getCODCONVENIO());

			Cell CNPJ = row.createCell(celula++);
			CNPJ.setCellValue(dado.getCNPJ());

			Cell MATRICULA = row.createCell(celula++);
			MATRICULA.setCellValue(dado.getMATRICULA());

			Cell CPF = row.createCell(celula++);
			CPF.setCellValue(dado.getCPF());

					
			Cell NOMEF = row.createCell(celula++);
			NOMEF.setCellValue(dado.getNOMEF());

			Cell LIMITEM = row.createCell(celula++);
			LIMITEM.setCellValue(dado.getLIMITEM());

			Cell DTNASCIMENTO = row.createCell(celula++);
			DTNASCIMENTO.setCellValue(dado.getDTNASCIMENTO());

			Cell CARGO = row.createCell(celula++);
			CARGO.setCellValue(dado.getCARGO());
			
			Cell DTEMISSAO = row.createCell(celula++);
			DTEMISSAO.setCellValue(dado.getDTEMISSAO());
			
			Cell NOMEMAE = row.createCell(celula++);
			NOMEMAE.setCellValue(dado.getNOMEMAE());

			Cell SUPIMEDIATO = row.createCell(celula++);
			SUPIMEDIATO.setCellValue(dado.getSUPIMEDIATO());
		
			Cell NACIONALIDADE = row.createCell(celula++);
			NACIONALIDADE.setCellValue(dado.getNACIONALIDADE());
	
			Cell ENDERECO = row.createCell(celula++);
			ENDERECO.setCellValue(dado.getENDERECO());
	
			Cell NUMERO = row.createCell(celula++);
			NUMERO.setCellValue(dado.getNUMERO());
	
			Cell COMPLEMENTO = row.createCell(celula++);
			COMPLEMENTO.setCellValue(dado.getCOMPLEMENTO());
			
			Cell BAIRRO = row.createCell(celula++);
			BAIRRO.setCellValue(dado.getBAIRRO());
			
			Cell CEP = row.createCell(celula++);
			CEP.setCellValue(dado.getCEP());
			
			Cell CIDADE = row.createCell(celula++);
			CIDADE.setCellValue(dado.getCIDADE());
			
			Cell ESTADO = row.createCell(celula++);
			ESTADO.setCellValue(dado.getESTADO());
		
			Cell DDD = row.createCell(celula++);
			DDD.setCellValue(dado.getDDD());
			
			Cell TELCELULAR = row.createCell(celula++);
			TELCELULAR.setCellValue(dado.getTELCELULAR());
			
			Cell EMAIL = row.createCell(celula++);
			EMAIL.setCellValue(dado.getEMAIL());
			
			Cell AFINSS = row.createCell(celula++);
			AFINSS.setCellValue(dado.getAFINSS());
			
			Cell AFMP = row.createCell(celula++);
			AFMP.setCellValue(dado.getAFMP());
            
			Cell SALBRUTO = row.createCell(celula++);
			SALBRUTO.setCellValue(dado.getSALBRUTO());
	
			Cell CODFILIAL = row.createCell(celula++);
			CODFILIAL.setCellValue(dado.getCODFILIAL());
	
			Cell NOMEFILILA = row.createCell(celula++);
			NOMEFILILA.setCellValue(dado.getNOMEFILILA());
	
			Cell CODCCUSTO = row.createCell(celula++);
			CODCCUSTO.setCellValue(dado.getCODCCUSTO());
		
			Cell NOMECCUSTO = row.createCell(celula++);
			NOMECCUSTO.setCellValue(dado.getNOMECCUSTO());
			
	
		}

		try {

			FileOutputStream out = new FileOutputStream(new File(nomeArquivo));
			workbook.write(out);
			out.close();
			System.out.println("*******************************");
			System.out.println("ARQUIVO CRIADO COM SUCESSO!!!");
			System.out.println("*******************************");
			System.out.println(nomeArquivo);

		} catch (Exception e) {
			System.out.println("ERRO!");
		}
		workbook.close();
	}

}
