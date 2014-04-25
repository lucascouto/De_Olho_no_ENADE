package parser;

import java.util.List;

public class Main {

	public static void main(String[] args) {
		List<String>valoresCelulas = UtilExcel.lerDadosPlanilha();

		
		System.out.println(valoresCelulas.get(0));
		System.out.println(valoresCelulas.get(1));
		System.out.println(valoresCelulas.get(2));
		System.out.println(valoresCelulas.get(3));
		System.out.println(valoresCelulas.get(19));
		System.out.println(valoresCelulas.get(20));
		System.out.println(valoresCelulas.get(21));
		System.out.println(valoresCelulas.get(22));
		System.out.println(valoresCelulas.get(23));
		
	}

}
