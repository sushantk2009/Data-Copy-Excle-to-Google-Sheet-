package org.example;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.*;

public class Main {

    private static final String APPLICATION_NAME = "Excel Sync";
    private static final String SPREADSHEET_ID = "1KeIWe1A8WaNLw7GWoaSPykxCIiRF82ZR298N20seEBw";
    private static final String RANGE = "Sheet1!A1";

    public static void main(String[] args) {
        try {

            // 🔑 Load credentials
            GoogleCredentials credentials = GoogleCredentials
                    .fromStream(new FileInputStream("credentials.json"))
                    .createScoped(Collections.singleton("https://www.googleapis.com/auth/spreadsheets"));

            // 🔗 Create Sheets service
            Sheets service = new Sheets.Builder(
                    GoogleNetHttpTransport.newTrustedTransport(),
                    GsonFactory.getDefaultInstance(),
                    new HttpCredentialsAdapter(credentials))
                    .setApplicationName(APPLICATION_NAME)
                    .build();

            // 📊 Read Excel file
            FileInputStream file = new FileInputStream("data.xlsx");
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            List<List<Object>> data = new ArrayList<>();

            int lastColumn = sheet.getRow(0).getLastCellNum();

            for (Row row : sheet) {
                List<Object> rowData = new ArrayList<>();

                for (int i = 0; i < lastColumn; i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    switch (cell.getCellType()) {
                        case STRING:
                            rowData.add(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            rowData.add(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            rowData.add(cell.getBooleanCellValue());
                            break;
                        default:
                            rowData.add("");
                    }
                }

                data.add(rowData);
            }

            workbook.close();

            // 📤 Upload to Google Sheets
            ValueRange body = new ValueRange().setValues(data);

            service.spreadsheets().values()
                    .update(SPREADSHEET_ID, RANGE, body)
                    .setValueInputOption("RAW")
                    .execute();

            System.out.println("✅ Data successfully uploaded!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}