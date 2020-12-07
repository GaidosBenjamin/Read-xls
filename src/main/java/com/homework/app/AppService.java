package com.homework.app;

import com.homework.entity.CustomerPurchase;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
public class AppService {

    public static void readExcel() {
        try {
            String excelFilePath = "sample.xls";
            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileInputStream);
            HSSFSheet sheet = hssfWorkbook.getSheetAt(0);
            DataFormatter dataFormatter = new DataFormatter();
            for(Row row : sheet) {
                for(Cell cell : row) {
                    if(cell.getColumnIndex() == 0) {
                        System.out.print(cell.getNumericCellValue() + " -- ");
                    } else {
                        System.out.print(dataFormatter.formatCellValue(cell) + " -- ");
                    }
                }
                System.out.println();
            }
            System.out.println();

            List<CustomerPurchase> list = new ArrayList<>();
            for(Row row : sheet) {
                var customerPurchase = new CustomerPurchase();
                for(Cell cell : row) {
                    switch(cell.getColumnIndex()) {
                        case 0:
                            customerPurchase.setId((int)cell.getNumericCellValue());
                            break;
                        case 1:
                            customerPurchase.setPurchasedProduct(dataFormatter.formatCellValue(cell));
                            break;
                        case 2:
                            customerPurchase.setName(dataFormatter.formatCellValue(cell));
                            break;
                        case 8:
                            customerPurchase.setCategory(dataFormatter.formatCellValue(cell));
                            break;
                    }

                }
                list.add(customerPurchase);
            }

            list.forEach(System.out::println);

            hssfWorkbook.close();
        }catch (Exception ignored) {}

    }


}
