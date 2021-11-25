import java.io.FileNotFoundException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Excute {

	public static void main(String[] args) throws FileNotFoundException {

		String plantilla = "C:\\test\\movimientos.xlsx";
		LocalDate fecha = LocalDate.now();

		List<String> moveList = new ArrayList<>();
		moveList.add(fecha.toString());
		moveList.add("First move");
		moveList.add("500");
		moveList.add("500");

		Map<String, Object> parameters = new HashMap<>();
		parameters.put("TITLES", "This is a test");
		parameters.put("MOVES", moveList);
		parameters.put("TOTAL", "111.111.111");
		parameters.put("ACCOUNT", "C.A " + " **** " + " 7584");
		parameters.put("CURRENCY_AMOUNT", "AMOUNT IN " + " USD");
		parameters.put("CURRECY_BALANCE", "AMOUNT IN " + "USD");

		Excel.downloadMoves(plantilla, parameters);

	}

}
