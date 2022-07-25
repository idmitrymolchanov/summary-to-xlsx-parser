package ru.example.service.impl;

import static ru.example.utils.HeaderEnum.DATE;
import static ru.example.utils.HeaderEnum.DISC_MODEL;
import static ru.example.utils.HeaderEnum.DISC_SERIAL_NUMBER;
import static ru.example.utils.HeaderEnum.DISC_TOTAL_SIZE;
import static ru.example.utils.HeaderEnum.HOST_ID;
import static ru.example.utils.HeaderEnum.HOURS_IN_OPERATION;
import static ru.example.utils.HeaderEnum.MB_CHECK;
import static ru.example.utils.HeaderEnum.MB_READ;
import static ru.example.utils.HeaderEnum.RESTART_CYCLE;
import static ru.example.utils.HeaderEnum.SIZE_IN_BYTES;
import static ru.example.utils.HeaderEnum.TEMPERATURE;
import static ru.example.utils.ProjectConstants.EXCEL_FILE_NAME;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.example.exception.FileNotFoundException;
import ru.example.exception.WriteWorkbookException;
import ru.example.model.HeaderModel;
import ru.example.service.ExcelService;

public class ExcelServiceImpl extends ExcelService {

  /**
   * @param inputDirPath  - путь до директории с файлами .summary
   * @param outputDirPath - путь до директории с конечным файлом .xlsx
   */
  @Override
  public void createWorkbook(String inputDirPath, String outputDirPath) {
    try (Stream<Path> files = Files.list(
        Paths.get(inputDirPath)); Workbook workbook = new XSSFWorkbook()) {

      final CellStyle headerStyle = getHeaderStyle(workbook);
      final CellStyle commonStyle = getCommonStyle(workbook);
      files.forEach(file -> createSheet(workbook, file.toString(), headerStyle, commonStyle));
      writeWorkbookToFile(outputDirPath, workbook);

    } catch (IOException e) {
      throw new FileNotFoundException(inputDirPath);
    }
  }

  private void writeWorkbookToFile(String outputDirPath, Workbook workbook) {
    try (FileOutputStream outputStream = new FileOutputStream(String.format("%s/%s.xlsx", outputDirPath, EXCEL_FILE_NAME))) {
      workbook.write(outputStream);
    } catch (IOException e) {
      throw new WriteWorkbookException(outputDirPath);
    }
  }


  /**
   *
   * @param workbook
   * @param pathString
   * @param headerStyle
   * @param commonStyle
   */
  private void createSheet(Workbook workbook, String pathString, CellStyle headerStyle, CellStyle commonStyle) {
    Path path = Paths.get(pathString);
    String fileName = path.getFileName().toString().replaceAll(".{8}$", "");
    Sheet sheet = workbook.createSheet(fileName);

    fillHeader(sheet, headerStyle);
    fillSheet(commonStyle, sheet, path);

    sheet.setAutoFilter(new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(), sheet.getLeftCol(), 11));
  }

  /**
   *
   * @param sheet
   * @param style
   */
  private void fillHeader(Sheet sheet, CellStyle style) {
    Row header = sheet.createRow(0);

    List<HeaderModel> headerCols = List.of(
        new HeaderModel(DATE.getKey(), 7000),
        new HeaderModel(HOST_ID.getKey(), 5000),
        new HeaderModel(DISC_MODEL.getKey(), 6000),
        new HeaderModel(DISC_SERIAL_NUMBER.getKey(), 6000),
        new HeaderModel(DISC_TOTAL_SIZE.getKey(), 7000),
        new HeaderModel(HOURS_IN_OPERATION.getKey(), 3000),
        new HeaderModel(RESTART_CYCLE.getKey(), 3000),
        new HeaderModel(MB_CHECK.getKey(), 3000),
        new HeaderModel(MB_READ.getKey(), 3000),
        new HeaderModel(SIZE_IN_BYTES.getKey(), 5000),
        new HeaderModel(TEMPERATURE.getKey(), 2000)
    );

    headerCols.forEach(column -> {
      int index = headerCols.indexOf(column);
      sheet.setColumnWidth(index, column.getWidth());
      createHeaderCell(header, column.getName(), index, style);
    });
  }

  /**
   * @param style
   * @param sheet
   * @param path
   */
  private void fillSheet(CellStyle style, Sheet sheet, Path path) {
    try (Stream<String> stream = Files.lines(path)) {
      List<String> list = stream.collect(Collectors.toList());
      int lineIndex = 1;
      for (String line : list) {
        addRecord(style, sheet, lineIndex++, line.split(";"));
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }


  /**
   * @param style
   * @param sheet
   * @param indexRow
   * @param values
   */
  private void addRecord(CellStyle style, Sheet sheet, int indexRow, String[] values) {

    Row row = sheet.createRow(indexRow);
    for (int i = 0; i < values.length; i++) {
      Cell cell = row.createCell(i);
      cell.setCellValue(values[i]);
      cell.setCellStyle(style);
    }
  }

  /**
   *
   * @param header
   * @param name
   * @param index
   * @param style
   */
  private void createHeaderCell(Row header, String name, int index, CellStyle style) {
    Cell headerCell = header.createCell(index);
    headerCell.setCellValue(name);
    headerCell.setCellStyle(style);
  }

}
