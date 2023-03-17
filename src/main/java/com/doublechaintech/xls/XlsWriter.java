package com.doublechaintech.xls;

import cn.hutool.core.codec.Base64;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import org.apache.poi.common.Duplicatable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.*;
import java.util.Iterator;
import java.util.List;

public class XlsWriter implements BlockWriter {
  private Workbook workBook;
  private boolean autoHeight = false;

  public XlsWriter(String base64) {
    if (ObjectUtil.isEmpty(base64)) {
      workBook = cn.hutool.poi.excel.WorkbookUtil.createBook(true);
      return;
    }
    byte[] decode = Base64.decode(base64);
    try {
      InputStream stream = new ByteArrayInputStream(decode);
      workBook = WorkbookFactory.create(stream);
      autoHeight = true;
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  public XlsWriter(File templateFile) {
    try {
      workBook = WorkbookFactory.create(new FileInputStream(templateFile));
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

    Block styleReferBlock = pBl.getStyleReferBlock();

    // 样式引用
    if (styleReferBlock != null) {
      Cell styleCell = ensureCell(styleReferBlock);
      CellStyle cellStyle = styleCell.getCellStyle();
      if (cellStyle instanceof Duplicatable) {
        Duplicatable copy = ((Duplicatable) cellStyle).copy();
        cell.setCellStyle((CellStyle) copy);
      } else {
        cell.getCellStyle().cloneStyleFrom(cellStyle);
      }
    }

    if (autoHeight) {
      cell.getCellStyle().setWrapText(true);
      cell.getRow().setHeight((short) -1);
    }

    if (pBl.getProperties() != null) {
      Number fillPattern = (Number) pBl.getProperties().get("fillPattern");
      if (fillPattern != null) {
        pBl.getProperties().put("fillPattern", fillPattern.shortValue());
      }
      CellUtil.setCellStyleProperties(cell, pBl.getProperties());
    }
  }

  @Override
  public void write(OutputStream out) throws IOException {
    int numberOfSheets = workBook.getNumberOfSheets();
    for (int i = 0; i < numberOfSheets; i++) {
      Sheet sheet = workBook.getSheetAt(i);
      sheet.setPrintGridlines(true);
      sheet.setDisplayGridlines(true);
    }

    Sheet sheetAt = workBook.getSheetAt(0);
    Iterator<Row> iterator = sheetAt.iterator();
    Row r = null;
    int maxColumns = -1;
    while (iterator.hasNext()) {
      r = iterator.next();

      Iterator<Cell> cellIterator = r.iterator();
      Cell cell = null;
      while (cellIterator.hasNext()) {
        cell = cellIterator.next();
      }

      if (cell != null && maxColumns < cell.getColumnIndex()) {
        maxColumns = cell.getColumnIndex();
      }
    }
    workBook.setPrintArea(0, 0, maxColumns, 0, r.getRowNum());
    workBook.write(out);
  }

  protected void setCellValue(Cell cell, Block pBlock) {
    cell.setCellValue(String.valueOf(pBlock.getValue()));
  }

  private Cell ensureCell(Block pBlock) {
    Row row = ensureRow(pBlock);
    Cell cell = row.getCell(pBlock.getLeft());
    if (cell == null) {
      cell = row.createCell(pBlock.getLeft());
    }
    return cell;
  }

  private Row ensureRow(Block pBlock) {
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
    }
    return row;
  }

  private boolean regionExisted(Sheet pSheet, CellRangeAddress pRegion) {
    List<CellRangeAddress> tMergedRegions = pSheet.getMergedRegions();
    if (tMergedRegions == null) {
      return false;
    }

    return tMergedRegions.contains(pRegion);
  }

  private Sheet ensureSheet(Block pBlock) {

    if (workBook == null) {
      throw new IllegalStateException("Work book is NOT READY yet!");
    }

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

  public static void main(String[] args) throws Exception {
    XlsWriter writer = new XlsWriter(new File("/Users/jackytian/Desktop/xls测试模板.xlsx"));

    Block style = new Block();
    style.setTop(0);
    style.setBottom(0);
    style.setLeft(0);
    style.setRight(0);

    Block data = new Block();
    data.setTop(1);
    data.setBottom(1);
    data.setLeft(1);
    data.setRight(1);
    data.setValue(StrUtil.repeat("一", 50));
    data.setStyleReferBlock(style);
    writer.append(data);

    data = new Block();
    data.setTop(2);
    data.setBottom(2);
    data.setLeft(2);
    data.setRight(3);
    data.setValue(StrUtil.repeat("一", 30));
    data.setStyleReferBlock(style);
    writer.append(data);

    writer.write(new FileOutputStream("/Users/jackytian/Desktop/xls测试模板-输出.xlsx"));
  }
}
