package parser;

//Classes Java necessárias para Leitura de arquivos e Iteração de Dados
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.io.FileInputStream;

//Libs POI para ler .xlsx
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UtilExcel {

	//Armazenar os valores de todas as celulas da planilha;
	private static List<String>valorCelulas = new ArrayList<String>();
	
	public static List<String> lerDadosPlanilha() {
	
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
				linhaIterator.next();
				linhaIterator.next();
				
				while(linhaIterator.hasNext()) {
					
					Row linha = linhaIterator.next();
					
					//Iterando sobre as colunas de uma linha;
					Iterator<Cell>celulaIterator = linha.cellIterator();
					while(celulaIterator.hasNext()) {
						
						Cell celula = celulaIterator.next();
						
						//Verificando o tipo da celula e dando o tratamento adequado;
						if(celula.getCellType() == Cell.CELL_TYPE_STRING) 
							valorCelulas.add(celula.getStringCellValue());
							
						
					}//Fim do while().
					
				}//Fim do while().
				
			}//Fim do for();
			
			//Fechando o FileInputStream fis;
			fis.close();
			
		}catch(IOException e) {
			e.printStackTrace();
		}
		
		return valorCelulas;
		
	}//Fim do lerDadosPlanilha().
	
}//Fim da classe;
