package parser;

public class Main {
	
	public static void main(String[] args) {
		
		System.out.println("Executando parser...");
		
		if(ParserExcel.lerDadosPlanilha() == true) {
			if(ParserExcel.gerarArquivoSQL() == true)
				System.out.println("Parser concluido.");
		}
		
		
	}

}
