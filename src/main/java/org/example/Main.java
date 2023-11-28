package org.example;

import lombok.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.println("How many books do you want: ");

        int numberOfBooks = scanner.nextInt();

        List<Book> books = new ArrayList<>();
        for (int i = 0; i < numberOfBooks; i++) {
            System.out.println("Enter details for Book " + (i + 1));
            System.out.print("Title: ");
            String title = scanner.next();

            System.out.print("Description: ");
            String description = scanner.next();

            System.out.print("Price: ");
            double price = scanner.nextDouble();

            System.out.print("Count: ");
            int count = scanner.nextInt();

            books.add(Book.builder()
                    .title(title)
                    .description(description)
                    .price(price)
                    .count(count)
                    .build());
        }

        bookToExcel(books);
    }

    private static void bookToExcel(List<Book> books) {
        try (FileOutputStream fileOut = new FileOutputStream("BookInformation.xlsx")) {
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Book Data");

                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Title");
                headerRow.createCell(1).setCellValue("Description");
                headerRow.createCell(2).setCellValue("Price");
                headerRow.createCell(3).setCellValue("Count");

                int rowNum = 1;
                for (Book book : books) {
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(book.getTitle());
                    row.createCell(1).setCellValue(book.getDescription());
                    row.createCell(2).setCellValue(book.getPrice());
                    row.createCell(3).setCellValue(book.getCount());
                }

                workbook.write(fileOut);
                System.out.println("Excel file created");

            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
