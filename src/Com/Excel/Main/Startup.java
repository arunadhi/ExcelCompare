package Com.Excel.Main;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

import Com.Excel.Compare.ReadExcelFile;

public class Startup {
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileReader reader = new FileReader("confing.properties");

		Properties p = new Properties();
		p.load(reader);

		String fileName = p.getProperty("filepath");
		String compareSheet1 = p.getProperty("compareSheet1");
		String compareSheet2 = p.getProperty("compareSheet2");
		ReadExcelFile.resultFilepath = p.getProperty("ResultPath");

		ReadExcelFile.readExceldata(fileName, compareSheet1, compareSheet2);
	}

}
