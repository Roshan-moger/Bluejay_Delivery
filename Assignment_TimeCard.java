package com.Assignment.Check;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.*;

public class App {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\user\\Downloads\\Assignment_Timecard.xlsx";

        // Step 1: Employees who worked consecutively for seven days
        printConsecutiveWorkingDays(filePath);
        
        System.out.println();
        // Step 2: Employees who have less than 10 hours of time between shifts but greater than 1 hour
        printShortBreakEmployees(filePath);
        System.out.println();

        // Step 3: Employees who worked for more than 14 hours in a single shift
        printLongShiftEmployees(filePath);
        System.out.println();
    }

    private static void printConsecutiveWorkingDays(String filePath) {
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(new File(filePath)))) {
            Sheet sheet = workbook.getSheetAt(0);
            System.out.println();

            System.out.println(" A:------> Employees who worked consecutively for seven days:");
            System.out.println("1------------------------------------------------------1");

            Map<String, Set<LocalDateTime>> employeeDates = new HashMap<>();
            String currentEmployee = null;

            for (int index = 0; index <= sheet.getLastRowNum(); index++) {
                Row row = sheet.getRow(index);
                LocalDateTime localDateTime = getLocalDateTimeCellValue(row, 2); // Assuming Column 3 is in the third column
                String positionId = getCellValue(row, 0); // Assuming Position ID is in the first column

                if (localDateTime != null && positionId != null) {
                    if (currentEmployee == null || !currentEmployee.equals(positionId)) {
                        currentEmployee = positionId;
                        employeeDates.put(currentEmployee, new HashSet<>());
                    }

                    Set<LocalDateTime> dates = employeeDates.get(currentEmployee);
                    dates.add(localDateTime);

                    if (dates.size() == 7) {
                        System.out.println("     Employee Name: " + getCellValue(row, 7));
                        // Assuming Employee Name is in the eighth column
                        System.out.println("1------------------------------------------------------1");
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void printShortBreakEmployees(String filePath) {
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(new File(filePath)))) {
            Sheet sheet = workbook.getSheetAt(0);
            System.out.println();

            System.out.println(" B:------>Name of Employees who have less than 10 hours of time between shifts but greater than 1 hour:");
            System.out.println("1-------------------------------------------------1");

            Map<String, LocalDateTime> employeeBreaks = new HashMap<>();
            Set<String> shortBreakPrinted = new HashSet<>();

            for (int index = 0; index <= sheet.getLastRowNum(); index++) {
                Row row = sheet.getRow(index);
                String employeeName = getCellValue(row, 7); // Assuming Employee Name is in the eighth column

                if (employeeName != null && !shortBreakPrinted.contains(employeeName)) {
                    LocalDateTime lastTimeOut = employeeBreaks.get(employeeName);
                    LocalDateTime timeIn = getLocalDateTimeCellValue(row, 2); // Assuming Time is in the third column

                    if (timeIn != null && lastTimeOut != null) {
                        double timeDiffHours = calculateTimeDifference(lastTimeOut, timeIn);

                        if (1 < timeDiffHours && timeDiffHours < 10) {
                            System.out.println("  Employee Name: " + employeeName);
                            shortBreakPrinted.add(employeeName);
                            System.out.println("1-------------------------------------------------1");
                        }
                    }

                    // Update the last time out for the employee
                    employeeBreaks.put(employeeName, getLocalDateTimeCellValue(row, 3)); // Assuming Time Out is in the fourth column
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void printLongShiftEmployees(String filePath) {
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(new File(filePath)))) {
            Sheet sheet = workbook.getSheetAt(0);
            	System.out.println();
            System.out.println(" C:------> Name Employees who worked for more than 14 hours:");
           

            Map<String, LocalDateTime> lastShiftEnd = new HashMap<>();
            Set<String> longShiftPrinted = new HashSet<>();

            for (int index = 0; index <= sheet.getLastRowNum(); index++) {
                Row row = sheet.getRow(index);
                String employeeName = getCellValue(row, 7); // Assuming Employee Name is in the eighth column
                String positionId = getCellValue(row, 0); // Assuming Position ID is in the first column

                if (employeeName != null && !longShiftPrinted.contains(employeeName)) {
                    LocalDateTime timeIn = getLocalDateTimeCellValue(row, 2); // Assuming Time is in the third column
                    LocalDateTime timeOut = getLocalDateTimeCellValue(row, 3); // Assuming Time Out is in the fourth column

                    if (timeIn != null && timeOut != null) {
                        double shiftDuration = calculateTimeDifference(timeIn, timeOut);

                        if (shiftDuration > 14) {
                            System.out.println("1---------------------------------------------------------------------1");
                            System.out.println("  Employee Name: " + employeeName + ", Position: " + positionId);
                            longShiftPrinted.add(employeeName);
                            System.out.println("1---------------------------------------------------------------------1");
                        }
                    }

                    // Update the last shift end time for the employee
                    lastShiftEnd.put(employeeName, timeOut);
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static LocalDateTime getLocalDateTimeCellValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
            double excelValue = cell.getNumericCellValue();
            return LocalDateTime.ofInstant(
                    java.time.Instant.ofEpochMilli((long) ((excelValue - 25569) * 86400 * 1000)),
                    java.time.ZoneId.systemDefault()
            );
        }
        return null;
    }

    private static String getCellValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        return (cell != null && cell.getCellType() == CellType.STRING) ? cell.getStringCellValue() : null;
    }

    private static double calculateTimeDifference(LocalDateTime startTime, LocalDateTime endTime) {
        long seconds = java.time.Duration.between(startTime, endTime).getSeconds();
        return seconds / 3600.0; // Convert seconds to hours
    }
}
