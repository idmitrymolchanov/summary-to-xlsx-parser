package ru.example.service;

import static ru.example.utils.HeaderEnum.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.example.model.HeaderModel;

public class ExcelService {

  private static final String HEADER_FONT_NAME = "Arial";
  private static final short HEADER_FONT_HEIGHT = 11;
  private static final boolean HEADER_IS_BOLD = true;
  private static final String EXCEL_FILE_NAME = "file";

  /**
   *
   * @throws IOException
   */
  public void createWorkbook(String inputDirPath, String outputDirPath) throws IOException {
    try (Stream<Path> files = Files.list(Paths.get(inputDirPath));
        Workbook workbook = new XSSFWorkbook()) {

      files.forEach(file -> {
        createSheet(workbook, String.valueOf(file.toAbsolutePath()));
        String fileLocation = String.format("%s/%s.xlsx", outputDirPath, EXCEL_FILE_NAME);
        FileOutputStream outputStream;
        try {
          outputStream = new FileOutputStream(fileLocation);
          workbook.write(outputStream);
        } catch (IOException e) {
          e.printStackTrace();
        }
      });

    }
  }

  /**
   *
   * @param workbook
   * @param pathString
   */
  public void createSheet(Workbook workbook, String pathString) {
    Path path = Paths.get(pathString);
    String fileName = String.valueOf(path.getFileName()).replaceAll(".{8}$", "");
    Sheet sheet = workbook.createSheet(fileName);

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
      createHeader(workbook, header, column.getName(), index);
    });

    CellStyle style = workbook.createCellStyle();
    style.setWrapText(true);

    fillSheet(style, sheet, path);

    sheet.setAutoFilter(new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(), sheet.getLeftCol(), 11));
  }


  /**
   *
   * @param style
   * @param sheet
   * @param path
   */
  private void fillSheet(CellStyle style, Sheet sheet, Path path) {
    try (Stream<String> stream = Files.lines(path)) {
      List<String> list = stream.collect(Collectors.toList());
      int lineIndex = 1;
      for(String line : list) addRecord(style, sheet, lineIndex++, line.split(";"));
    } catch (IOException e) {
      e.printStackTrace();
    }
  }


  /**
   *
   * @param style
   * @param sheet
   * @param indexRow
   * @param values
   */
  private void addRecord (CellStyle style, Sheet sheet, int indexRow, String[] values) {

    Row row = sheet.createRow(indexRow);
    for(int i = 0; i < values.length; i++) {
      Cell cell = row.createCell(i);
      cell.setCellValue(values[i]);
      cell.setCellStyle(style);
    }
    //    sheet.setAutoFilter(new CellRangeAddress(cell.getRow().getFirstCellNum(), cell.getRow().getLastCellNum(), sheet.getFirstRowNum(), sheet.getLastRowNum()));
  }

  /**
   *
   * @param workbook
   * @param header
   * @param name
   * @param index
   */
  private static void createHeader(Workbook workbook, Row header, String name, int index) {
    Cell headerCell = header.createCell(index);
    headerCell.setCellValue(name);
    headerCell.setCellStyle(getHeaderStyle(workbook));
  }

  /**
   *
   * @param workbook
   * @return
   */
  private static CellStyle getHeaderStyle(Workbook workbook) {
    CellStyle headerStyle = workbook.createCellStyle();
    headerStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    XSSFFont font = ((XSSFWorkbook) workbook).createFont();
    font.setFontName(HEADER_FONT_NAME);
    font.setFontHeightInPoints(HEADER_FONT_HEIGHT);
    font.setBold(HEADER_IS_BOLD);
    font.setColor(HSSFColorPredefined.WHITE.getIndex());
    headerStyle.setFont(font);
    return headerStyle;
  }

}