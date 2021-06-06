package domaciPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BrojeviSuma {

	public static void main(String[] args) {
	//Napisati program koji racuna sumu brojeva koji se nalaze prvom sheet-u 
	//excel tabele koji se zove Brojevi. 
	//U tabeli svi brojevi se nalaze u prvoj koloni. 
	//Program treba da cita red po red iz tabele i upisane brojeve dodaje na sumu.
	//Ukupnu sumu na kraju treba ispisati na standardnom izlazu. 
	//Potrebno je omoguciti da program radi i ukoliko se u datu tabelu doda jos brojeva.

	File f = new File("DomaciMoj.xlsx");
	double zbir = 0;

	try {
		InputStream is = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(is);
		Sheet brojevi = wb.getSheetAt(0);
		// System.out.println(brojevi.getLastRowNum());

		for (int i = 0; i < brojevi.getLastRowNum() + 1; i++) {

			Row row = brojevi.getRow(i);
			Cell cell = row.getCell(0);
			double broj = cell.getNumericCellValue();
			zbir = zbir + broj;
			// System.out.println(broj + " zbir ta 2 broja je: " + zbir);

		}
		// System.out.println(zbir);

		wb.close();

	} catch (NullPointerException | IOException e) {

		e.printStackTrace();

	} finally {
		System.out.println("Konacan zbir je: " + zbir);

	}

}

}
