package br.com.bulla.LeitorPlanilha;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Dados {

	
	//classe formatação de celulas 
		public static String formataDados03(String dado){
			   dado = dado
					    .replace("(","")
					    .replace(")","")
					    .replace("/","")
					    .replace(":","")
					    .replace(">","")
					    .replace("<","")
					    .replace(",","")
					    .replace("]","")
					    .replace("}","")
					    .replace("[","")
					    .replace("{","")
					    .replace("","")
					    .replace("+","")
					    .replace("=","")
					    .replace("_","")
					    .replace("@","")
					    .replace("!","")
					    .replace("?","")
					    .replace("\\","")
					    .replace("|","")
					    .replace("~","")
					    .replace("^","")
					    .replace(";","")
					    .replace(".","")
					    .replace("ç", "c")   
		                .replace("Ç", "C")   
		                .replace("ñ", "n")   
		                .replace("Ñ", "N")
		                .replace("#", "")
		                .replace("$", "")
		                .replace("%", "")
		                .replace("'", "")
		                .replace("&", "")
		                .replace("\"", "")
		                .replace("*", "")
		                .replace("ã", "a")
		                .replace("õ", "o")
		                .replaceAll("[êèéë]", "e")   
		                .replaceAll("[îìíï]", "i")   
		                .replaceAll("[õôòóö]", "o")   
		                .replaceAll("[ûúùü]", "u")   
		                .replaceAll("[ÃÂÀÁÄ]", "A")   
		                .replaceAll("[ÊÈÉË]", "E")   
		                .replaceAll("[ÎÌÍÏ]", "I")   
		                .replaceAll("[ÕÔÒÓÖ]", "O")   
		                .replaceAll("[ÛÙÚÜ]", "U")
		                .replaceAll("[a-z]", "")
		                .replaceAll("[A-Z]", "")
		                ;
			   
			// 
			   
			   return dado;
			}

	
	
	
	//classe formatação de celulas 
	public static String formataDados06(String dado){
		   dado = dado
				    .replace("(","")
				    .replace(")","")
				    .replace("/","")
				    .replace(":","")
				    .replace(">","")
				    .replace("<","")
				    .replace(",","")
				    .replace("]","")
				    .replace("}","")
				    .replace("[","")
				    .replace("{","")
				    .replace("","")
				    .replace("+","")
				    .replace("-","")
				    .replace("=","")
				    .replace("_","")
				    .replace("@","")
				    .replace("!","")
				    .replace("?","")
				    .replace("\\","")
				    .replace("|","")
				    .replace("~","")
				    .replace("^","")
				    .replace(";","")
				    .replace(".","")
				    .replace("ç", "c")   
	                .replace("Ç", "C")   
	                .replace("ñ", "n")   
	                .replace("Ñ", "N")
	                .replace("#", "")
	                .replace("$", "")
	                .replace("%", "")
	                .replace("'", "")
	                .replace("&", "")
	                .replace("\"", "")
	                .replace("*", "")
	                .replace("ã", "a")
	                .replace("õ", "o")
	                .replaceAll("[êèéë]", "e")   
	                .replaceAll("[îìíï]", "i")   
	                .replaceAll("[õôòóö]", "o")   
	                .replaceAll("[ûúùü]", "u")   
	                .replaceAll("[ÃÂÀÁÄ]", "A")   
	                .replaceAll("[ÊÈÉË]", "E")   
	                .replaceAll("[ÎÌÍÏ]", "I")   
	                .replaceAll("[ÕÔÒÓÖ]", "O")   
	                .replaceAll("[ÛÙÚÜ]", "U")
	                .replaceAll("[1234567890]", "")
	                ;
		   
		// 
		   
		   return dado;
		}

	//classe formatação de celulas 
		public static String formataDados07(String dado){
			   dado = dado
					    .replace("(","")
					    .replace(")","")
					    .replace("/","")
					    .replace(":","")
					    .replace(">","")
					    .replace("<","")
					    .replace("]","")
					    .replace("}","")
					    .replace("[","")
					    .replace("{","")
					    .replace("","")
					    .replace("+","")
					    .replace("-","")
					    .replace("=","")
					    .replace("_","")
					    .replace("@","")
					    .replace("!","")
					    .replace("?","")
					    .replace("\\","")
					    .replace("|","")
					    .replace("~","")
					    .replace("^","")
					    .replace(";","")
					    .replace(".","")
					    .replace("ç", "c")   
		                .replace("Ç", "C")   
		                .replace("ñ", "n")   
		                .replace("Ñ", "N")
		                .replace("#", "")
		                .replace("$", "")
		                .replace("%", "")
		                .replace("'", "")
		                .replace("&", "")
		                .replace("\"", "")
		                .replace("*", "")
		                .replace("ã", "a")
		                .replace("õ", "o")
		                .replace("¨", "")
		                .replaceAll("[êèéë]", "e")   
		                .replaceAll("[îìíï]", "i")   
		                .replaceAll("[õôòóö]", "o")   
		                .replaceAll("[ûúùü]", "u")   
		                .replaceAll("[ÃÂÀÁÄ]", "A")   
		                .replaceAll("[ÊÈÉË]", "E")   
		                .replaceAll("[ÎÌÍÏ]", "I")   
		                .replaceAll("[ÕÔÒÓÖ]", "O")   
		                .replaceAll("[ÛÙÚÜ]", "U")
		                .replaceAll("[a-z]", "")
		                .replaceAll("[A-Z]", "")
		                ;
			   
			// 
			   
			   return dado;
			}

	
		//classe formatação de celulas 
		public static String formataDados08(String dado){
			   dado = dado
					    .replace("(","")
					    .replace(")","")
					    .replace(":","")
					    .replace(">","")
					    .replace("<","")
					    .replace(",","")
					    .replace("]","")
					    .replace("}","")
					    .replace("[","")
					    .replace("{","")
					    .replace("","")
					    .replace("+","")
					    .replace("-","")
					    .replace("=","")
					    .replace("_","")
					    .replace("@","")
					    .replace("!","")
					    .replace("?","")
					    .replace("\\","")
					    .replace("|","")
					    .replace("~","")
					    .replace("^","")
					    .replace(";","")
					    .replace(".","")
					    .replace("ç", "c")   
		                .replace("Ç", "C")   
		                .replace("ñ", "n")   
		                .replace("Ñ", "N")
		                .replace("#", "")
		                .replace("$", "")
		                .replace("%", "")
		                .replace("'", "")
		                .replace("&", "")
		                .replace("\"", "")
		                .replace("*", "")
		                .replace("ã", "a")
		                .replace("õ", "o")
		                .replace("¨", "")
		                .replaceAll("[êèéë]", "e")   
		                .replaceAll("[îìíï]", "i")   
		                .replaceAll("[õôòóö]", "o")   
		                .replaceAll("[ûúùü]", "u")   
		                .replaceAll("[ÃÂÀÁÄ]", "A")   
		                .replaceAll("[ÊÈÉË]", "E")   
		                .replaceAll("[ÎÌÍÏ]", "I")   
		                .replaceAll("[ÕÔÒÓÖ]", "O")   
		                .replaceAll("[ÛÙÚÜ]", "U")
		                .replaceAll("[a-z]", "")
		                .replaceAll("[A-Z]", "")
		                ;
			   
			// 
			   
			   return dado;
			}
	

		//classe formatação de celulas 
		public static String formataDados18(String dado){
			   dado = dado
					    .replace("(","")
					    .replace(")","")
					    .replace(":","")
					    .replace(">","")
					    .replace("<","")
					    .replace(",","")
					    .replace("]","")
					    .replace("}","")
					    .replace("[","")
					    .replace("{","")
					    .replace("","")
					    .replace("+","")
					     .replace("=","")
					    .replace("_","")
					    .replace("@","")
					    .replace("!","")
					    .replace("?","")
					    .replace("\\","")
					    .replace("|","")
					    .replace("~","")
					    .replace("^","")
					    .replace(";","")
					    .replace(".","")
					    .replace("ç", "c")   
		                .replace("Ç", "C")   
		                .replace("ñ", "n")   
		                .replace("Ñ", "N")
		                .replace("#", "")
		                .replace("$", "")
		                .replace("%", "")
		                .replace("'", "")
		                .replace("&", "")
		                .replace("\"", "")
		                .replace("*", "")
		                .replace("ã", "a")
		                .replace("õ", "o")
		                .replace("¨", "")
		                .replaceAll("[êèéë]", "e")   
		                .replaceAll("[îìíï]", "i")   
		                .replaceAll("[õôòóö]", "o")   
		                .replaceAll("[ûúùü]", "u")   
		                .replaceAll("[ÃÂÀÁÄ]", "A")   
		                .replaceAll("[ÊÈÉË]", "E")   
		                .replaceAll("[ÎÌÍÏ]", "I")   
		                .replaceAll("[ÕÔÒÓÖ]", "O")   
		                .replaceAll("[ÛÙÚÜ]", "U")
		                .replaceAll("[a-z]", "")
		                .replaceAll("[A-Z]", "")
		                ;
			   
			// 
			   
			   return dado;
			}
	
		//classe formatação de celulas 
		public static String formataDados21(String dado){
			   dado = dado
					    .replace("(","")
					    .replace(")","")
					    .replace("/","")
					    .replace(":","")
					    .replace(">","")
					    .replace("<","")
					    .replace(",","")
					    .replace("]","")
					    .replace("}","")
					    .replace("[","")
					    .replace("{","")
					    .replace("","")
					    .replace("+","")
					    .replace("-","")
					    .replace("=","")
					    .replace("_","")
					    .replace("@","")
					    .replace("!","")
					    .replace("?","")
					    .replace("\\","")
					    .replace("|","")
					    .replace("~","")
					    .replace("^","")
					    .replace(";","")
					    .replace(".","")
					    .replace("ç", "c")   
		                .replace("Ç", "C")   
		                .replace("ñ", "n")   
		                .replace("Ñ", "N")
		                .replace("#", "")
		                .replace("$", "")
		                .replace("%", "")
		                .replace("'", "")
		                .replace("&", "")
		                .replace("\"", "")
		                .replace("*", "")
		                .replace("ã", "a")
		                .replace("õ", "o")
		                .replace(" ", "")
		                .replace("¨", "")
		                .replaceAll("[êèéë]", "e")   
		                .replaceAll("[îìíï]", "i")   
		                .replaceAll("[õôòóö]", "o")   
		                .replaceAll("[ûúùü]", "u")   
		                .replaceAll("[ÃÂÀÁÄ]", "A")   
		                .replaceAll("[ÊÈÉË]", "E")   
		                .replaceAll("[ÎÌÍÏ]", "I")   
		                .replaceAll("[ÕÔÒÓÖ]", "O")   
		                .replaceAll("[ÛÙÚÜ]", "U")
		                .replaceAll("[a-z]", "")
		                .replaceAll("[A-Z]", "")
		                ;
			   
			// 
			   
			   return dado;
			}

	
		//classe formatação de celulas 
		public static String formataDados23(String dado){
			   dado = dado
					    .replace("(","")
					    .replace(")","")
					    .replace("/","")
					    .replace(":","")
					    .replace(">","")
					    .replace("<","")
					    .replace(",","")
					    .replace("]","")
					    .replace("}","")
					    .replace("[","")
					    .replace("{","")
					    .replace("","")
					    .replace("+","")
					    .replace("=","")
					    .replace("_","")
					    .replace("!","")
					    .replace("?","")
					    .replace("\\","")
					    .replace("|","")
					    .replace("~","")
					    .replace("^","")
					    .replace(";","")
					    .replace(".","")
					    .replace("ç", "c")   
		                .replace("Ç", "C")   
		                .replace("ñ", "n")   
		                .replace("Ñ", "N")
		                .replace("#", "")
		                .replace("$", "")
		                .replace("%", "")
		                .replace("'", "")
		                .replace("&", "")
		                .replace("\"", "")
		                .replace("*", "")
		                .replace("ã", "a")
		                .replace("õ", "o")
		                .replace("¨", "")
		                .replaceAll("[êèéë]", "e")   
		                .replaceAll("[îìíï]", "i")   
		                .replaceAll("[õôòóö]", "o")   
		                .replaceAll("[ûúùü]", "u")   
		                .replaceAll("[ÃÂÀÁÄ]", "A")   
		                .replaceAll("[ÊÈÉË]", "E")   
		                .replaceAll("[ÎÌÍÏ]", "I")   
		                .replaceAll("[ÕÔÒÓÖ]", "O")   
		                .replaceAll("[ÛÙÚÜ]", "U")
		                 ;
			   
			// 
			   
			   return dado;
			}

		
		//classe formatação de celulas 
		public static String formataDados27(String dado){
			   dado = dado
					    .replace("(","")
					    .replace(")","")
					    .replace("/","")
					    .replace(":","")
					    .replace(">","")
					    .replace("<","")
					    .replace(",","")
					    .replace("]","")
					    .replace("}","")
					    .replace("[","")
					    .replace("{","")
					    .replace("","")
					    .replace("+","")
					    .replace("-","")
					    .replace("=","")
					    .replace("_","")
					    .replace("@","")
					    .replace("!","")
					    .replace("?","")
					    .replace("\\","")
					    .replace("|","")
					    .replace("~","")
					    .replace("^","")
					    .replace(";","")
					    .replace(".","")
					    .replace("ç", "c")   
		                .replace("Ç", "C")   
		                .replace("ñ", "n")   
		                .replace("Ñ", "N")
		                .replace("#", "")
		                .replace("$", "")
		                .replace("%", "")
		                .replace("'", "")
		                .replace("&", "")
		                .replace("\"", "")
		                .replace("*", "")
		                .replace("ã", "a")
		                .replace("õ", "o")
		                .replace("¨", "")
		                .replaceAll("[êèéë]", "e")   
		                .replaceAll("[îìíï]", "i")   
		                .replaceAll("[õôòóö]", "o")   
		                .replaceAll("[ûúùü]", "u")   
		                .replaceAll("[ÃÂÀÁÄ]", "A")   
		                .replaceAll("[ÊÈÉË]", "E")   
		                .replaceAll("[ÎÌÍÏ]", "I")   
		                .replaceAll("[ÕÔÒÓÖ]", "O")   
		                .replaceAll("[ÛÙÚÜ]", "U")
		                ;
			   
			// 
			   
			   return dado;
			}

		
		
		private String OPERACAO;
	private String CODCONVENIO;
	private String CNPJ;
	private String MATRICULA;
	private String CPF;
	private String NOMEF;
	private String LIMITEM;
	private String DTNASCIMENTO;
	private String CARGO;
	private String DTEMISSAO;
	private String NOMEMAE;
	private String SUPIMEDIATO;
	private String NACIONALIDADE;
	private String ENDERECO;
	private String NUMERO;
	private String COMPLEMENTO;
	private String BAIRRO;
	private String CEP;
	private String CIDADE;
	private String ESTADO;
	private String DDD;
	private String TELCELULAR;
	private String EMAIL;
	private String AFINSS;
	private String AFMP;
	private String SALBRUTO;
	private String CODFILIAL;
	private String NOMEFILILA;
	private String CODCCUSTO;
	private String NOMECCUSTO;

	public Dados(String OPERACAO, String CODCONVENIO, String CNPJ, String MATRICULA, String CPF, String NOMEF,
			String LIMITEM, String DTNASCIMENTO, String CARGO, String DTEMISSAO, String NOMEMAE, String SUPIMEDIATO,
			String NACIONALIDADE, String ENDERECO, String NUMERO, String COMPLEMENTO, String BAIRRO, String CEP,
			String CIDADE, String ESTADO, String DDD, String TELCELULAR, String EMAIL, String AFINSS, String AFMP,
			String SALBRUTO, String CODFILIAL, String NOMEFILILA, String CODCCUSTO, String NOMECCUSTO) {

		this.OPERACAO = OPERACAO;
		this.CODCONVENIO = CODCONVENIO;
		this.CNPJ = CNPJ;
		this.MATRICULA = MATRICULA;
		this.CPF = CPF;
		this.NOMEF = NOMEF;
		this.LIMITEM = LIMITEM;
		this.DTNASCIMENTO = DTNASCIMENTO;
		this.CARGO = CARGO;
		this.DTEMISSAO = DTEMISSAO;
		this.NOMEMAE = NOMEMAE;
		this.SUPIMEDIATO = SUPIMEDIATO;
		this.NACIONALIDADE = NACIONALIDADE;
		this.ENDERECO = ENDERECO;
		this.NUMERO = NUMERO;
		this.COMPLEMENTO = COMPLEMENTO;
		this.BAIRRO = BAIRRO;
		this.CEP = CEP;
		this.CIDADE = CIDADE;
		this.ESTADO = ESTADO;
		this.DDD = DDD;
		this.TELCELULAR = TELCELULAR;
		this.EMAIL = EMAIL;
		this.AFINSS = AFINSS;
		this.AFMP = AFMP;
		this.SALBRUTO = SALBRUTO;
		this.CODFILIAL = CODFILIAL;
		this.NOMEFILILA = NOMEFILILA;
		this.CODCCUSTO = CODCCUSTO;
		this.NOMECCUSTO = NOMECCUSTO;

	}

	public String getOPERACAO() {
		return this.OPERACAO;
	}

	public String getCODCONVENIO() {
		return this.CODCONVENIO;
	}

	public String getCNPJ() {
		return this.CNPJ;
	}

	public String getMATRICULA() {
		return this.MATRICULA;
	}

	public String getCPF() {
		return this.CPF;
	}

	public String getNOMEF() {
		return this.NOMEF;
	}

	public String getLIMITEM() {
		return this.LIMITEM;
	}

	public String getDTNASCIMENTO() {
		return this.DTNASCIMENTO;
	}

	public String getCARGO() {
		return this.CARGO;
	}

	public String getDTEMISSAO() {
		return this.DTEMISSAO;
	}

	public String getNOMEMAE() {
		return this.NOMEMAE;
	}

	public String getSUPIMEDIATO() {
		return this.SUPIMEDIATO;
	}

	public String getNACIONALIDADE() {
		return this.NACIONALIDADE;
	}

	public String getENDERECO() {
		return this.ENDERECO;
	}

	public String getNUMERO() {
		return this.NUMERO;
	}

	public String getCOMPLEMENTO() {
		return this.COMPLEMENTO;
	}

	public String getBAIRRO() {
		return this.BAIRRO;
	}

	public String getCEP() {
		return this.CEP;
	}

	public String getCIDADE() {
		return this.CIDADE;
	}

	public String getESTADO() {
		return this.ESTADO;
	}

	public String getDDD() {
		return this.DDD;
	}

	public String getTELCELULAR() {
		return this.TELCELULAR;
	}

	public String getEMAIL() {
		return this.EMAIL;
	}

	public String getAFINSS() {
		return this.AFINSS;
	}

	public String getAFMP() {
		return this.AFMP;
	}

	public String getSALBRUTO() {
		return this.SALBRUTO;
	}

	public String getCODFILIAL() {
		return this.CODFILIAL;
	}

	public String getNOMEFILILA() {
		return this.NOMEFILILA;
	}

	public String getCODCCUSTO() {
		return this.CODCCUSTO;
	}

	public String getNOMECCUSTO() {
		return this.NOMECCUSTO;
	}

	public static List<Dados> getDados() throws IOException {

		List<Dados> dados = new ArrayList<>();
		String filename = "C:\\LerPlanilha\\planilha.xlsx";
		FileInputStream stream = new FileInputStream(filename);

		XSSFWorkbook workbook = new XSSFWorkbook(stream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.iterator();

		while (rowIterator.hasNext()) {

			Row row = rowIterator.next();
			if (row.getRowNum() == 0)
				continue;
		
				
				
			Iterator<Cell> cellIterator = row.cellIterator();

			/* COLUNA - 01 */String OPERACAO = "";
			/* COLUNA - 02 */String CODCONVENIO = "";
			/* COLUNA - 03 */String CNPJ = "";
			/* COLUNA - 04 */String MATRICULA = "";
			/* COLUNA - 05 */String CPF = "";
			/* COLUNA - 06 */String NOMEF = "";
			/* COLUNA - 07 */String LIMITEM = "";
			/* COLUNA - 08 */String DTNASCIMENTO = "";
			/* COLUNA - 09 */String CARGO = "";
			/* COLUNA - 10 */String DTEMISSAO = "";
			/* COLUNA - 11 */String NOMEMAE = "";
			/* COLUNA - 12 */String SUPIMEDIATO = "";
			/* COLUNA - 13 */String NACIONALIDADE = "";
			/* COLUNA - 14 */String ENDERECO = "";
			/* COLUNA - 15 */String NUMERO = "";
			/* COLUNA - 16 */String COMPLEMENTO = "";
			/* COLUNA - 17 */String BAIRRO = "";
			/* COLUNA - 18 */String CEP = "";
			/* COLUNA - 19 */String CIDADE = "";
			/* COLUNA - 20 */String ESTADO = "";
			/* COLUNA - 21 */String DDD = "";
			/* COLUNA - 22 */String TELCELULAR = "";
			/* COLUNA - 23 */String EMAIL = "";
			/* COLUNA - 24 */String AFINSS = "";
			/* COLUNA - 25 */String AFMP = "";
			/* COLUNA - 26 */String SALBRUTO = "";
			/* COLUNA - 27 */String CODFILIAL = "";
			/* COLUNA - 28 */String NOMEFILILA = "";
			/* COLUNA - 29 */String CODCCUSTO = "";
			/* COLUNA - 30 */String NOMECCUSTO = "";

			while (cellIterator.hasNext()) {

				DataFormatter formatar = new DataFormatter();

				Cell cell = cellIterator.next();
				switch (cell.getColumnIndex()) {

				case 0:
					// nao faz nada

				case 1:
					OPERACAO = formatar.formatCellValue(cell);

				case 2:
					CODCONVENIO = formatar.formatCellValue(cell);

				case 3:
					CNPJ =  formataDados03(formatar.formatCellValue(cell));

				case 4:
					MATRICULA = formatar.formatCellValue(cell);

				case 5:
					CPF = formatar.formatCellValue(cell);

				case 6:
					NOMEF = formataDados06(formatar.formatCellValue(cell));

				case 7:
					LIMITEM = formataDados07(formatar.formatCellValue(cell));

				case 8:
					DTNASCIMENTO = formataDados08(formatar.formatCellValue(cell));

				case 9:
					CARGO = formatar.formatCellValue(cell);

				case 10:
					DTEMISSAO = formataDados08(formatar.formatCellValue(cell));

				case 11:
					NOMEMAE = formataDados06(formatar.formatCellValue(cell));

				case 12:
					SUPIMEDIATO = formataDados06(formatar.formatCellValue(cell));

				case 13:
					NACIONALIDADE = formatar.formatCellValue(cell);

				case 14:
					ENDERECO = formataDados06(formatar.formatCellValue(cell));

				case 15:
					NUMERO = formatar.formatCellValue(cell);

				case 16:
					COMPLEMENTO = formatar.formatCellValue(cell);

				case 17:
					BAIRRO = formatar.formatCellValue(cell);

				case 18:
					CEP = formataDados18(formatar.formatCellValue(cell));

				case 19:
					CIDADE = formatar.formatCellValue(cell);

				case 20:
					ESTADO = formatar.formatCellValue(cell);

				case 21:
					DDD = formataDados21(formatar.formatCellValue(cell));

				case 22:
					TELCELULAR = formataDados18(formatar.formatCellValue(cell));

				case 23:
					EMAIL = formataDados18(formatar.formatCellValue(cell));

				case 24:
					AFINSS = formatar.formatCellValue(cell);

				case 25:
					AFMP = formatar.formatCellValue(cell);

				case 26:
					SALBRUTO = formataDados07(formatar.formatCellValue(cell));

				case 27:
					CODFILIAL = formataDados27(formatar.formatCellValue(cell));

				case 28:
					NOMEFILILA = formataDados27(formatar.formatCellValue(cell));

				case 29:
					CODCCUSTO = formataDados27(formatar.formatCellValue(cell));

				case 30:
					NOMECCUSTO = formataDados27(formatar.formatCellValue(cell));

					break;
				}

			}

			dados.add(new Dados(OPERACAO, CODCONVENIO, CNPJ, MATRICULA, CPF, NOMEF, LIMITEM, DTNASCIMENTO, CARGO,
					DTEMISSAO, NOMEMAE, SUPIMEDIATO, NACIONALIDADE, ENDERECO, NUMERO, COMPLEMENTO, BAIRRO, CEP, CIDADE,
					ESTADO, DDD, TELCELULAR, EMAIL, AFINSS, AFMP, SALBRUTO, CODFILIAL, NOMEFILILA, CODCCUSTO,
					NOMECCUSTO));
		}

		workbook.close();
		stream.close();

		return dados;

	}

}
