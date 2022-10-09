package com.doublechaintech.xls;

import cn.hutool.core.codec.Base64;
import cn.hutool.core.util.ObjectUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class XlsWriter implements BlockWriter {
  private Workbook workBook;

  public XlsWriter(String base64) {
    if (ObjectUtil.isEmpty(base64)) {
      workBook = cn.hutool.poi.excel.WorkbookUtil.createBook(true);
      return;
    }
    byte[] decode = Base64.decode(base64);
    try {
      InputStream stream = new ByteArrayInputStream(decode);
      workBook = WorkbookFactory.create(stream);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  @Override
  public void append(Block pBl) {
    if (pBl == null) {
      return;
    }
    Cell cell = ensureCell(pBl);
    setCellValue(cell, pBl);
    if (pBl.getProperties() != null) {
      CellUtil.setCellStyleProperties(cell, pBl.getProperties());
    }
  }

  @Override
  public void write(OutputStream out) throws IOException {
    workBook.write(out);
  }

  public CellStyle getDefaultStyle() {
    final XSSFCellStyle style = (XSSFCellStyle) workBook.createCellStyle();
    style.setBorderBottom(BorderStyle.THIN);
    return style;
  }

  protected void setCellValue(Cell cell, Block pBlock) {
    cell.setCellStyle(getDefaultStyle());
    cell.setCellValue(String.valueOf(pBlock.getValue()));
  }

  private Cell ensureCell(Block pBlock) {
    Sheet sheet = ensureSheet(pBlock);

    // region, let's create the region
    if (!isCell(pBlock)) {
      CellRangeAddress region =
          new CellRangeAddress(
              pBlock.getTop(), pBlock.getBottom(), pBlock.getLeft(), pBlock.getRight());

      if (!regionExisted(sheet, region)) {

        // here also may fail for intersection, will do nothing here
        sheet.addMergedRegion(region);
      }
    }

    // use the primary cell/ left-top
    Row row = sheet.getRow(pBlock.getTop());
    if (row == null) {
      row = sheet.createRow(pBlock.getTop());
      row.setHeight((short) 500);
    }
    Cell cell = row.getCell(pBlock.getLeft());
    if (cell == null) {
      cell = row.createCell(pBlock.getLeft());
    }
    return cell;
  }

  private boolean regionExisted(Sheet pSheet, CellRangeAddress pRegion) {
    List<CellRangeAddress> tMergedRegions = pSheet.getMergedRegions();
    if (tMergedRegions == null) {
      return false;
    }

    return tMergedRegions.contains(pRegion);
  }

  private Sheet ensureSheet(Block pBlock) {
    String page = pBlock.getPage();

    // the sheet
    String sheetName = sheetName(page);
    Sheet sheet = null;
    if (sheetName == null) {
      // no sheet name, we will try first sheet
      if (workBook.getNumberOfSheets() > 0) {
        sheet = workBook.getSheetAt(0);
      }
      if (sheet == null) {
        sheet = workBook.createSheet();
      }
    } else {
      sheet = workBook.getSheet(sheetName);
      if (sheet == null) {
        sheet = workBook.createSheet(sheetName);
      }
    }
    return sheet;
  }

  private String sheetName(String pPage) {
    if (pPage == null) {
      return null;
    }
    return WorkbookUtil.createSafeSheetName(pPage);
  }

  private boolean isCell(Block pBlock) {
    return pBlock.getBottom() == pBlock.getTop() && pBlock.getLeft() == pBlock.getRight();
  }
}
