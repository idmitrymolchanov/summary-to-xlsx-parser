package ru.example;

import java.io.IOException;
import ru.example.service.ExcelService;

public class Application {

  public static void main(String[] args) throws IOException {
    ExcelService excelService = new ExcelService();
    excelService.createWorkbook("/Users/dmolchanov/Desktop/Новая папка/test", "/Users/dmolchanov/Desktop/Новая папка");
  }


}
