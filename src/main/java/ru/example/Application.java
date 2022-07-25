package ru.example;

import javax.swing.JFileChooser;
import ru.example.service.ExcelService;
import ru.example.service.impl.ExcelServiceImpl;

public class Application {

  public static void main(String[] args) {
    JFileChooser f = new JFileChooser();
    f.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
    f.setDialogTitle("Выберите директорию с файлами .summary");
    f.showOpenDialog(null);

    ExcelService excelService = new ExcelServiceImpl();
    excelService.createWorkbook(f.getSelectedFile().getPath(), "/Users/dmolchanov/Desktop/Новая папка");
  }

}