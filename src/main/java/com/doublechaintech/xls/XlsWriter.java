package com.doublechaintech.xls;

import cn.hutool.core.codec.Base64;
import cn.hutool.core.util.ObjectUtil;
import org.apache.poi.common.Duplicatable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.*;
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

      // 处理高度
      String styleCellValue = styleCell.getStringCellValue();
      String currentValue = cell.getStringCellValue();
      if (ObjectUtil.isNotEmpty(styleCellValue) && ObjectUtil.isNotEmpty(currentValue)) {
        int styleValueLength = styleCellValue.length();
        int cellValueLength = currentValue.length();
        int styleHeight = getHeight(styleReferBlock);
        int styleColumnWidth = getWidth(styleReferBlock);
        int cellColumnWidth = getWidth(pBl);
        int lines =
            (styleColumnWidth * cellValueLength + cellColumnWidth * styleValueLength - 1)
                / (cellColumnWidth * styleValueLength);
        if (lines > 1) {
          // 高度，自动换行
          short currentHeight = cell.getRow().getHeight();
          short requiredHeight = (short) (styleHeight * lines);
          if (currentHeight < requiredHeight) {
            cell.getRow().setHeight(requiredHeight);
          }
          cell.getCellStyle().setWrapText(true);
        }
      }
    }

    if (pBl.getProperties() != null) {
      Number fillPattern = (Number) pBl.getProperties().get("fillPattern");
      if (fillPattern != null) {
        pBl.getProperties().put("fillPattern", fillPattern.shortValue());
      }
      CellUtil.setCellStyleProperties(cell, pBl.getProperties());
    }
  }

  // 获取一个块的宽度
  public int getWidth(Block block) {
    int width = 0;
    Sheet sheet = ensureSheet(block);
    for (int i = block.getLeft(); i <= block.getRight(); i++) {
      width += sheet.getColumnWidth(i);
    }
    return width;
  }

  // 获取一个块的高度
  public int getHeight(Block block) {
    int height = 0;
    Sheet sheet = ensureSheet(block);
    for (int i = block.getTop(); i <= block.getBottom(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) {
        row = sheet.createRow(i);
      }
      height += row.getHeight();
    }
    return height;
  }

  @Override
  public void write(OutputStream out) throws IOException {
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
}
