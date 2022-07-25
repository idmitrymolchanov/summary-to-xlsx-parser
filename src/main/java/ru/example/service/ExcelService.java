package ru.example.service;

import static ru.example.utils.ProjectConstants.*;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public abstract class ExcelService {

  public abstract void createWorkbook(String inputDirPath, String outputDirPath);

  /**
   * @param workbook
   * @return
   */
  protected CellStyle getHeaderStyle(Workbook workbook) {
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

  /**
   *
   * @param workbook
   * @return
   */
  protected CellStyle getCommonStyle(Workbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setWrapText(true);
    return style;
  }

}