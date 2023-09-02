
package app;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class App {
    public static void main(String[] args) {

        System.out.println("- Caso o arquivo não exista na memória, será gerado um novo! -");
        System.out.println("Copie aqui o path do seu arquivo excel (com a extensão .xlsx): ");

        String path = new Scanner(System.in).next();

        File file = new File(path);

        //Nomes mokados
        Map<String, Integer> map = Map.of("Pedro", 10, "Jorge", 20);

        if (!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException e) {
                System.out.println("Error! " + e.getMessage());
            }
        }

        //Nova pasta de trabalho Excel
        try (Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("Test");

            //Retorna a última linha criada na planilha e vai p/ próxima
            int actualLine = createDefaultFields(sheet) + 1;
            for (Map.Entry<String, Integer> person : map.entrySet()) {

                String name = person.getKey();
                Integer age = person.getValue();

                Row linePerson = sheet.createRow(actualLine++);
                int cells = 0;

                Cell cellName = linePerson.createCell(cells++);
                cellName.setCellValue(name);

                Cell cellAge = linePerson.createCell(cells);
                cellAge.setCellValue(age);
            }

            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            outputStream.flush();

            outputStream.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }

    }


    private static int createDefaultFields(Sheet sheet) {

        Row row = sheet.createRow(0);

        Cell cellName = row.createCell(0);
        cellName.setCellValue("NAMES ");

        Cell cellAge = row.createCell(1);
        cellAge.setCellValue("AGES ");

        return sheet.getLastRowNum();
    }

}
