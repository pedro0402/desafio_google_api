import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

public class SpreadsheetReader {
    private static final String SPREADSHEET_ID = "1MJ1JOGTRBGpb41jZ57gj_ndE2lS2tWjeA1XlyDosibk";
    private static final String RANGE = "B4:F";

    public static void main(String[] args) throws IOException, GeneralSecurityException {
        Sheets service = SheetsQuickStart.getSheetsService();
        ValueRange response = service.spreadsheets().values().get(SPREADSHEET_ID, RANGE).execute();
        List<List<Object>> values = response.getValues();
        if (values == null || values.isEmpty()) {
            System.out.println("No data found.");
        } else {
            for (List<Object> row : values) {
                System.out.printf(" %s %s %s %s %s\n", row.get(0), row.get(1), row.get(2), row.get(3), row.get(4));
            }

            List<Double> averageGrades = calculateAverageGrades(values);
            List<Integer> studentAbsences = extractAbsences(values);
            List<String> studentStatuses = new ArrayList<>();
            System.out.println("Average grades and student statuses:");
            for (int i = 0; i < averageGrades.size(); i++) {
                double average = averageGrades.get(i);
                int absences = studentAbsences.get(i);
                String status = determineStatus(average, absences);
                studentStatuses.add(status);
                System.out.printf("Student %d: Average = %.2f, Absences = %d, Status = %s\n", i + 1, average, absences,
                        studentStatuses.get(i));
            }

            writeStatusesToSpreadsheet(service, values, studentStatuses, averageGrades);
        }
    }

    private static List<Double> calculateAverageGrades(List<List<Object>> values) {
        List<Double> averages = new ArrayList<>();
        for (List<Object> row : values) {
            int totalGrades = 0;
            double sumGrades = 0;
            for (int i = 2; i < row.size(); i++) {
                try {
                    double grade = Double.parseDouble(row.get(i).toString());
                    totalGrades++;
                    sumGrades += grade;
                } catch (NumberFormatException e) {
                    System.out.println("Error converting grade: " + e.getMessage());
                }
            }
            if (totalGrades > 0) {
                double average = sumGrades / totalGrades;
                averages.add(average);
            } else {
                System.out.println("No valid grade found for a student.");
                averages.add(0.0);
            }
        }
        return averages;
    }

    private static List<Integer> extractAbsences(List<List<Object>> values) {
        List<Integer> absences = new ArrayList<>();
        for (List<Object> row : values) {
            try {
                int studentAbsences = Integer.parseInt(row.get(1).toString());
                absences.add(studentAbsences);
            } catch (NumberFormatException e) {
                System.out.println("Error converting absences: " + e.getMessage());
            }
        }
        return absences;
    }

    private static String determineStatus(double average, int absences) {
        if (absences > 12) {
            return "Failed due to Absences";
        }
        if (average < 70) {
            return "Final Exam";
        } else {
            return "Passed";
        }
    }

    private static void writeStatusesToSpreadsheet(Sheets service, List<List<Object>> values, List<String> statuses,
            List<Double> averageGrades) throws IOException {
        List<List<Object>> data = new ArrayList<>();
        for (int i = 0; i < statuses.size(); i++) {
            List<Object> rowData = new ArrayList<>();
            rowData.add(statuses.get(i));
            if (statuses.get(i).equals("Final Exam")) {
                double naf = calculateNAF(averageGrades.get(i));
                rowData.add(naf);
            } else {
                rowData.add(0);
            }
            data.add(rowData);
        }
        ValueRange body = new ValueRange().setValues(data);
        UpdateValuesResponse result = service.spreadsheets().values()
                .update(SPREADSHEET_ID, "G4:H" + (values.size() + 3), body).setValueInputOption("RAW").execute();
        System.out.printf("%d cells updated.", result.getUpdatedCells());
    }

    private static double calculateNAF(double average) {
        return average - 10;
    }
}
