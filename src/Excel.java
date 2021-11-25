
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static byte[] downloadMoves(String plantilla, Map<String, Object> parameters) throws FileNotFoundException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		FileOutputStream out = new FileOutputStream("C:\\test\\Report.xls");
		InputStream stream = new FileInputStream(plantilla);
		try (XSSFWorkbook documento = new XSSFWorkbook(stream)) {
			XSSFSheet hoja = documento.getSheet("Ultimos Movimientos");

			set(hoja, 3, 1, parameters.get("ACCOUNT").toString());
			set(hoja, 1, 9, parameters.get("TITLES").toString());
			set(hoja, 6, 14, parameters.get("CURRENCY_AMOUNT").toString());
			set(hoja, 6, 16, parameters.get("CURRECY_BALANCE").toString());

			Object moves = parameters.get("MOVES");
			List<Object> lstMovimientos =  (List<Object>) moves;
			
			Integer filaActual = 7;
				duplicarFila(hoja, filaActual);
				set(hoja, filaActual, 1, lstMovimientos.get(0).toString());
				set(hoja, filaActual, 3, lstMovimientos.get(1).toString());
				set(hoja, filaActual, 14, lstMovimientos.get(2).toString());
				set(hoja, filaActual, 16, lstMovimientos.get(3).toString());
				++filaActual;
			eliminarFila(hoja, filaActual++);

			set(hoja, filaActual, 14, parameters.get("TOTAL").toString());

			documento.write(baos);
			documento.write(out);
			
			
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		return baos.toByteArray();
	}

	private static void set(Sheet hoja, Integer fila, Integer columna, String valor) {
		Cell celda = hoja.getRow(fila).getCell(columna);
		if (celda == null) {
			celda = hoja.getRow(fila).createCell(columna);
		}
		celda.setCellValue(valor);
	}

	private static void duplicarFila(XSSFSheet hoja, Integer fila) {
		hoja.shiftRows(fila, hoja.getLastRowNum(), 1);
		hoja.copyRows(fila + 1, fila + 1, fila, new CellCopyPolicy());
	}

	private static void eliminarFila(XSSFSheet hoja, Integer fila) {
		hoja.removeRow(hoja.getRow(fila));
	}

}
