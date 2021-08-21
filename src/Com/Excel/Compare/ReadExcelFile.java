package Com.Excel.Compare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map.Entry;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {
	public static XSSFWorkbook compareBook;
	public static int getRowCountofSht1 = 0;
	public static int getRowCountofSht2 = 0;
	public static int getColCountOfsht1 = 0;
	public static int getColCountOfsht2 = 0;
	public static ArrayList<String> NotHavinTar = new ArrayList<>();
	public static ArrayList<String> NotHavinSrc = new ArrayList<>();
	public static String resultFilepath=null;

	public static void readExceldata(String fileName, String compareSheet1, String compareSheet2) throws IOException {

		File file = new File(fileName); // creating a new file instance
		FileInputStream fis = new FileInputStream(file);
		// creating Workbook instance that refers to .xlsx file
		compareBook = new XSSFWorkbook(fis);
		XSSFSheet compareSht1 = compareBook.getSheet(compareSheet1);
		XSSFSheet compareSht2 = compareBook.getSheet(compareSheet2);

		// get count of rows and column of both sheets
		getRowCountofSht1 = compareSht1.getPhysicalNumberOfRows();
		getRowCountofSht2 = compareSht2.getPhysicalNumberOfRows();
		getColCountOfsht1 = compareSht1.getRow(0).getPhysicalNumberOfCells();
		getColCountOfsht2 = compareSht2.getRow(0).getPhysicalNumberOfCells();

		CheckVaildForComparision(getColCountOfsht1, getColCountOfsht2, compareSht1, compareSht2);

		LinkedHashMap<String, ArrayList<String>> CompareMap1 = ExtractRowandColumn(getRowCountofSht1, getColCountOfsht1,
				compareSht1);
		LinkedHashMap<String, ArrayList<String>> CompareMap2 = ExtractRowandColumn(getRowCountofSht2, getColCountOfsht2,
				compareSht2);
		LinkedHashMap<String, ArrayList<Integer>> diffDataMap = compareSrcAndTargetSheet(CompareMap1, CompareMap2);
		compareSht1 = StyleDesinforDifferntvalueinSheet(diffDataMap, compareSht1, getRowCountofSht1);
		compareSht2 = StyleDesinforDifferntvalueinSheet(diffDataMap, compareSht2, getRowCountofSht2);
		FileOutputStream outputStream = new FileOutputStream(resultFilepath+"FinalOutput.xlsx"); 
			compareBook.write(outputStream);
			outputStream.close();
			compareBook.close();

	}

	public static LinkedHashMap<String, ArrayList<String>> ExtractRowandColumn(int rowCount, int colCount,
			XSSFSheet sheet) {
		LinkedHashMap<String, ArrayList<String>> ListRowOfwithColData = new LinkedHashMap<>();

		for (int i = 1; i < rowCount; i++) {
			ArrayList<String> arrycoldata = new ArrayList<>();
			XSSFRow rowData = sheet.getRow(i);

			for (int j = 0; j < colCount; j++) {

				XSSFCell coldata = rowData.getCell(j);

				arrycoldata.add(coldata.toString());

			}

			ListRowOfwithColData.put(rowData.getCell(0).toString(), arrycoldata);

		}

		return ListRowOfwithColData;

	}

	public static void CheckVaildForComparision(int srcColCount, int tarColCount, XSSFSheet Srcsheet,
			XSSFSheet tarsheet) {
		JFrame frame = new JFrame();

		if (srcColCount != tarColCount) {

			JOptionPane.showMessageDialog(frame, "Column Count Are Mismatched \n for both Sorce and target sheet",
					"Error Message Box", JOptionPane.ERROR_MESSAGE);
			System.exit(0);
		} else {
			XSSFRow srcRow = Srcsheet.getRow(0);
			XSSFRow tarRow = tarsheet.getRow(0);

			for (int i = 0; i < srcColCount; i++) {
				String srcColumnheader = srcRow.getCell(i).getStringCellValue();
				String tarColumnheader = tarRow.getCell(i).getStringCellValue();
				if (!srcColumnheader.equalsIgnoreCase(tarColumnheader)) {

					JOptionPane.showMessageDialog(frame,
							"Columns are mismatched or\n Column order was changed .\nSo we counld not compare the sheet",
							"Error Message Box", JOptionPane.ERROR_MESSAGE);
					System.exit(0);
				}

			}
		}

	}

	public static LinkedHashMap<String, ArrayList<Integer>> compareSrcAndTargetSheet(
			LinkedHashMap<String, ArrayList<String>> CompareMap1,
			LinkedHashMap<String, ArrayList<String>> CompareMap2) {

		NotHavinTar = new ArrayList<>();
		NotHavinSrc = new ArrayList<>();

		LinkedHashMap<String, ArrayList<Integer>> misMatchData = new LinkedHashMap<String, ArrayList<Integer>>();

		for (Entry<String, ArrayList<String>> entry : CompareMap1.entrySet()) {
			String srckey = entry.getKey();
			ArrayList<Integer> misMatchArry = new ArrayList<Integer>();
			ArrayList<String> srcvalue = entry.getValue();
			ArrayList<String> tarvalue = CompareMap2.get(srckey);

			if (CompareMap2.containsKey(srckey)) {

				for (int i = 0; i < srcvalue.size(); i++) {

					if (!srcvalue.get(i).equals(tarvalue.get(i))) {

						misMatchArry.add(i);

						// System.out.println(srcvalue.get(i) + " ---- "+ i + " --------- "+
						// tarvalue.get(i));

					}
				}

			} else {
				NotHavinTar.add(srckey);

			}
			misMatchData.put(srckey, misMatchArry);

		}

		for (Entry<String, ArrayList<String>> entry : CompareMap2.entrySet()) {
			String tarkey = entry.getKey();
			Object tarvalue = entry.getValue();

			if (!CompareMap1.containsKey(tarkey)) {

				NotHavinSrc.add(tarkey);

			}

		}
		return misMatchData;

	}

	public static XSSFSheet StyleDesinforDifferntvalueinSheet(LinkedHashMap<String, ArrayList<Integer>> DiffMap,
			XSSFSheet sheet, int rowCount) {

		XSSFSheet Stylesheet = sheet;
		for (Entry<String, ArrayList<Integer>> entry : DiffMap.entrySet()) {

			String failRow_Id = entry.getKey();

			for (int i = 1; i < rowCount; i++) {
				String UniValue = Stylesheet.getRow(i).getCell(0).toString();
				if (UniValue.equalsIgnoreCase(failRow_Id)) {
					for (int j = 0; j < entry.getValue().size(); j++) {
						XSSFCell cell = Stylesheet.getRow(i).getCell(entry.getValue().get(j));
						CellStyle style = compareBook.createCellStyle();
						// Setting Background color
						style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						cell.setCellStyle(style);

					}
				}

			}

		}
		return Stylesheet;

	}
}
