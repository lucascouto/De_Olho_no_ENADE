package parser;

//Classes Java necess�rias para Leitura de arquivos e Itera��o de Dados
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.PrintWriter;



//Libs POI para ler .xlsx
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ParserExcel {
	
	private static List<String>nomeCurso = new ArrayList<String>();//1
	private static List<Integer>codIES = new ArrayList<Integer>();//2
	private static List<String>nomeIES = new ArrayList<String>();//3
	private static List<String>tipoIES = new ArrayList<String>();//4
	private static List<String>orgAcade = new ArrayList<String>();//5
	private static List<String>municipio = new ArrayList<String>();//7
	private static List<String>uf = new ArrayList<String>();//9
	private static List<Integer>numEstudCurso = new ArrayList<Integer>();//11
	private static List<Integer>numEstudInsc = new ArrayList<Integer>();//12
	private static List<Float>conceitoEnade = new ArrayList<Float>();//17
	
	private static List<Integer>codIESAux = new ArrayList<Integer>();
	
	public static boolean lerDadosPlanilha() {
	
		try {
			
			FileInputStream fis = new FileInputStream("conceito_enade_2012.xlsx");
			
			//Variavel Workbook para xlsx FileInputStream;
			Workbook workbook = new XSSFWorkbook(fis);
			
			//Numero de planilhas no arquivo .xlsx;
			int numPlanilhas = workbook.getNumberOfSheets();
			
			//Loop sobre todas as planilhas do arquivo .xlsx;
			for(int i = 0; i < numPlanilhas; i++) {
				
				Sheet planilha = workbook.getSheetAt(i);
				
				//Iterando sobre as linhas de uma planilha;
				Iterator<Row> linhaIterator = planilha.iterator();
				
				//Come�ando da linha 3 da planilha;
				linhaIterator.next();
				linhaIterator.next();
				
				while(linhaIterator.hasNext()) {
					
					Row linha = linhaIterator.next();
					
					int posicaoColuna = 0;//Auxiliar para saber num da coluna;
					
					//Iterando sobre as colunas de uma linha;
					Iterator<Cell>celulaIterator = linha.cellIterator();
					while(celulaIterator.hasNext()) {
						
						Cell celula = celulaIterator.next();
						
						//Posicao coluna == 1, 2, 3, 4, 5, 7, 9, 11,12, 17
						switch(posicaoColuna) {
						
						case 1:
							if(celula.getCellType() == Cell.CELL_TYPE_STRING)
								nomeCurso.add(celula.getStringCellValue());
							else
								nomeCurso.add(".");
							break;
						case 2:
							codIES.add((int)celula.getNumericCellValue());
							break;
						case 3:
							if(celula.getCellType() == Cell.CELL_TYPE_STRING)
								nomeIES.add(celula.getStringCellValue());
							else
								nomeIES.add(".");
							break;
						case 4:
							if(celula.getCellType() == Cell.CELL_TYPE_STRING)
								tipoIES.add(celula.getStringCellValue());
							else
								tipoIES.add(".");
							break;
						case 5:
							if(celula.getCellType() == Cell.CELL_TYPE_STRING)
								orgAcade.add(celula.getStringCellValue());
							else
								orgAcade.add(".");
							break;
						case 7:
							if(celula.getCellType() == Cell.CELL_TYPE_STRING)
								municipio.add(celula.getStringCellValue());
							else
								municipio.add(".");
							break;
						case 9:
							if(celula.getCellType() == Cell.CELL_TYPE_STRING)
								uf.add(celula.getStringCellValue());
							else
								uf.add(".");
							break;
						case 11:
							if(celula.getCellType() == Cell.CELL_TYPE_NUMERIC)
								numEstudCurso.add((int)celula.getNumericCellValue());
							else
								numEstudCurso.add(0);
							break;
						case 12:
							if(celula.getCellType() == Cell.CELL_TYPE_NUMERIC)
								numEstudInsc.add((int)celula.getNumericCellValue());
							else
								numEstudInsc.add(0);
							break;
						case 17:
							if(celula.getCellType() == Cell.CELL_TYPE_NUMERIC)
								conceitoEnade.add((float)celula.getNumericCellValue());
							else
								conceitoEnade.add((float)0);
							break;
						default:
							break;
						}//Fim do switch().
						posicaoColuna++;
					}//Fim do while().
					
				}//Fim do while().
				
			}//Fim do for();
			
			//Fechando o FileInputStream fis;
			fis.close();
			
		}catch(IOException e) {
			e.printStackTrace();
			return false;
		}
		
		return true;
	}//Fim do lerDadosPlanilha().
	
	public static boolean gerarArquivoSQL() {
			
		String insertInstituicaoSQL = "INSERT INTO instituicao (cod_ies, org_academica, uf, "+
						   			"nome, tipo) VALUES (%d, '%s', '%s', '%s', '%s');"+"\n";
		
		try {
			PrintWriter writer = new PrintWriter("dados.sql");
			
			for(int i = 0; i < codIES.size(); i++) {
				//Escrevendo INSERT na tabela INSTITUICAO;
				
				int cod_ies = codIES.get(i);
				if(!codIESAux.contains(cod_ies)) {
					
					codIESAux.add(cod_ies);
					String org_academica = orgAcade.get(i);
					String ufIES = uf.get(i);
					String nome = nomeIES.get(i);
					if(nome.contains("'")) {
						String string[] = nome.split("'");
						nome = string[0]+string[1];
					}
					String tipo = tipoIES.get(i);
		
					writer.format(insertInstituicaoSQL, cod_ies, org_academica, ufIES, nome, tipo);
					
				}
			}
			
			writer.close();
			
			//Escrevendo INSERT na tabela CURSO;
			/*System.out.println("Tabela curso:");
			System.out.println("num_estud_curso: "+valoresInt.get(1));
			System.out.println("num_estud_insc: "+valoresInt.get(2));
			System.out.println("nome: "+valoresString.get(0));
			System.out.println("municipio: "+valoresString.get(4));
			System.out.println("conceito_enade: "+valoresFloat.get(0));*/
				
				
				
		}catch(FileNotFoundException e) {
			e.printStackTrace();
			return false;
		}catch(SecurityException e) {
			e.printStackTrace();
			return false;
		}
			
		return true;

	}//Fim do gerarArquivoSQL();
	
}//Fim da classe;
