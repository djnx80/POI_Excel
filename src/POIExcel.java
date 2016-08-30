import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileOutputStream;

public class POIExcel {

	public static void main(String[] args) {
			// stw�rz nowy zeszyt
			Workbook workbook = new HSSFWorkbook();
			// i dwie zaak�adki
			Sheet sheet1 = workbook.createSheet("Zak�adka 1");
			Sheet sheet2 = workbook.createSheet("Zak�adka 2");
			// w pierwszej zak�adce stw�rz nowe kom�rki
			Cell cell1 = sheet1.createRow(0).createCell(0);
			Cell cell2 = sheet1.createRow(0).createCell(2);
			Cell cell3 = sheet1.createRow(0).createCell(4);
			// i przypisz im warto�ci
			cell1.setCellValue(20);
			cell2.setCellValue(30);
			cell3.setCellFormula("SUM(A1:D1)");
			// w drugiej zak�adce wpisz tekst
			cell1 = sheet2.createRow(0).createCell(0);
			cell1.setCellValue("tekst w drugiej zak�adce");
			
			// odczytaj warto�� (cell1.getRichStringCellValue().toString())
		
			try {
				// je�eli mo�liwe to stw�rz nowy plik i zapisz tam warto�ci zeszytu
				FileOutputStream plik = new FileOutputStream("proba.xls");
				workbook.write(plik);
				// na koniec zamknij wszystko
				workbook.close();
				plik.close();
			}catch	(Exception e){
				e.printStackTrace();
			}

	}

}
