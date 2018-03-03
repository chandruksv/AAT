package com.etrade.aat.AAT;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;

public class ReadExcelFile {
	public static final String filePath = "C:" + File.separator + "AAT" + File.separator + "TestScenarios"
			+ File.separator + "TestScenarios.xlsx";
	public static final String templatePath = "C:" + File.separator + "AAT" + File.separator + "FeatureTemplates";
	public static final String featurePath = "C:" + File.separator + "AAT" + File.separator + "TestScripts";

	public static void main(String[] args) throws IOException {

		System.out.println(filePath);

		FileInputStream inputstream = new FileInputStream(new File(filePath));
		XSSFWorkbook wb = new XSSFWorkbook(inputstream);
		XSSFSheet ws = wb.getSheetAt(0);
		int n = ws.getLastRowNum();
		// System.out.println(n);

		XSSFRow row = null;
		XSSFCell desccell, idcell, pbicell, projectcell = null;
		String testID, pbi, userame = null;

		for (int i = 0; i <= n; i++) {
			row = ws.getRow(i);

			desccell = row.getCell(7);
			idcell = row.getCell(1);
			pbicell = row.getCell(2);
			projectcell = row.getCell(5);
			String designer = System.getProperty("user.name");
			// System.out.println(designer);
			String scenario = desccell.getStringCellValue().toLowerCase();
			testID = idcell.getStringCellValue();
			if (!testID.equals(null)) {
				pbi = pbicell.getStringCellValue();
				// updateFiles(pbi, designer);
				if (!(scenario.contains("scenario") | (scenario.contains("description")))) {
					// System.out.println(scenario);
					if (scenario.contains("account")
							&& (scenario.contains("import") && (scenario.contains("success")))) {
						updateFiles(pbi, designer, "Account_Import_Success.feature");
						createFeatureFiles(testID, "Account_Import_Success.feature");
					}
					else if (scenario.contains("exercise")
							&& (scenario.contains("import") && (scenario.contains("success")))) {
						updateFiles(pbi, designer, "Exercise_Import_Success.feature");
						createFeatureFiles(testID, "Exercise_Import_Success.feature");
					}
					else if (scenario.contains("exercise")
							&& (scenario.contains("import") && (scenario.contains("fail")))) {
						updateFiles(pbi, designer, "Exercise_Import_Failed.feature");
						createFeatureFiles(testID, "Exercise_Import_Failed.feature");
					}
					else if (scenario.contains("release")
							&& (scenario.contains("import") && (scenario.contains("fail")))) {
						updateFiles(pbi, designer, "Release_Import_Failed.feature");
						createFeatureFiles(testID, "Release_Import_Failed.feature");
					}
					else if (scenario.contains("release")
							&& (scenario.contains("import") && (scenario.contains("success")))) {
						updateFiles(pbi, designer, "Release_Import_Success.feature");
						createFeatureFiles(testID, "Release_Import_Success.feature");
					}
					else System.out.println("Automation Test Script Template not found for Scenario : "+testID);
				}
			}
		}

	}

	public static void updateFiles(String pbi, String designer, String file) throws IOException {
		File fileName = new File(templatePath + File.separator + file);
		ArrayList<String> lines = new ArrayList<String>();

		BufferedReader br = new BufferedReader(new FileReader(fileName));
		String fileContent;
		while ((fileContent = br.readLine()) != null) {
			if (fileContent.startsWith("#User Story:")) {
				fileContent = "#User Story:" + pbi;
			} else if (fileContent.startsWith("#Designer:")) {
				fileContent = "#Designer:" + designer;
			}
			lines.add(fileContent);
			fileContent = null;
		}
		br.close();
		BufferedWriter bw = new BufferedWriter(new FileWriter(fileName));
		for (String line : lines) {
			bw.write(line);
			bw.newLine();
		}
		bw.close();
		System.out.println("File" + file + "updated successfully");

	}

	public static void createFeatureFiles(String testID, String file) throws IOException {
		File sourceFile = new File(templatePath + File.separator + file);
		File destinationFile = new File(featurePath + File.separator + testID + ".feature");
		InputStream is = new FileInputStream(sourceFile);
		OutputStream os = new FileOutputStream(destinationFile);
		byte[] buffer = new byte[1024];
        int length;
        while ((length = is.read(buffer)) > 0) {
            os.write(buffer, 0, length);
        }
        is.close();
        os.close();

	}

}
