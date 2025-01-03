/*
 * Copyright 2003-2005 ExcelUtils http://excelutils.sourceforge.net
 * Created on 2005-6-22
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License. 
 */
package io.github.xuhai19901018.excelutil.tags;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.StringTokenizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import io.github.xuhai19901018.excelutil.ExcelParser;
import io.github.xuhai19901018.excelutil.ExcelUtils;
import io.github.xuhai19901018.excelutil.WorkbookUtils;



/**
 * <p>
 * <b>ForeachTag </b> is a class which parse the #foreach tag
 * </p>
 * 
 * @author rainsoft
 * @version $Revision: 1.6 $ $Date: 2005/07/12 08:25:57 $
 */
public class ForeachTag implements ITag {

  public static final String KEY_FOREACH = "#foreach";

  public static final String KEY_END = "#end";

  public int[] parseTag(Object context, Sheet sheet, Row curRow, Cell curCell) throws Exception {
    int forstart = curRow.getRowNum();
    int forend = -1;
    int forCount = 0;
    String foreach = "";
    boolean bFind = false;
    for (int rownum = forstart; rownum <= sheet.getLastRowNum(); rownum++) {
       if (rownum < 0) {
        continue;
       }
      Row row = sheet.getRow(rownum);
      if (null == row)
        continue;
      for (short colnum = row.getFirstCellNum(); colnum <= row.getLastCellNum(); colnum++) {
        if (colnum < 0) {
            continue;
        }
        Cell cell = row.getCell(colnum);
        if (null == cell)
          continue;
        if (cell.getCellType() == CellType.STRING) {
          String cellstr = cell.getStringCellValue();

          // get the tag instance for the cellstr
          ITag tag = ExcelParser.getTagClass(cellstr);

          if (null != tag) {
            if (tag.hasEndTag()) {
              if (0 == forCount) {
                forstart = rownum;
                foreach = cellstr;
              }
              forCount++;
              break;
            }
          }
          if (cellstr.startsWith(KEY_END)) {
            forend = rownum;
            forCount--;
            if (forstart >= 0 && forend >= 0 && forend > forstart && forCount == 0) {
              bFind = true;
            }
            break;
          }
        }
      }
      if (bFind)
        break;
    }

    if (!bFind)
      return new int[] { 0, 0, 1 };

    String properties = "";
    String property = "";
    // parse the collection an object
    StringTokenizer st = new StringTokenizer(foreach, " ");
    int pos = 0;
    while (st.hasMoreTokens()) {
      String str = st.nextToken();
      if (pos == 1) {
        property = str;
      }
      if (pos == 3) {
        properties = str;
      }
      pos++;
    }
    // get collection
    Object collection = ExcelParser.parseStr(context, properties);
    if (null == collection) {
      return new int[] { 0, 0, 1 };
    }
    // get the iterator of collection
    Iterator iterator = ExcelParser.getIterator(collection);

    // iterator
    int shiftNum = forend - forstart - 1;
    // set the start row number
    final int StartRowNo = forstart+1;
    ExcelUtils.addValue(context, property+"StartRowNo", StartRowNo);
    
    int old_forend = forend;
    int propertyId = 0;
    int shift = 0;
    if (null != iterator) {
      while (iterator.hasNext()) {
        Object obj = iterator.next();
        
        ExcelUtils.addValue(context, property, obj);
        // Iterator ID
        ExcelUtils.addValue(context, property + "Id", new Integer(propertyId));
        // Index start with 1
        ExcelUtils.addValue(context, property + "Index", new Integer(propertyId+1));
        
        // shift the #foreach #end block
        sheet.shiftRows(forstart, sheet.getLastRowNum(), shiftNum, true, true);
        // copy the body fo #foreach #end block
        WorkbookUtils.copyRow(sheet, forstart + shiftNum + 1, forstart, shiftNum);
        // parse
        shift = ExcelParser.parse(context, sheet, forstart, forstart + shiftNum - 1);

        forstart += shiftNum + shift;
        forend += shiftNum + shift;
        propertyId++;
      }
    }
    // set the end row number
    final int EndRowNo = forstart;
    ExcelUtils.addValue(context, property+"EndRowNo", forstart);

    // 2022年10月25日、合并掉循环的单元格
    List<CellRangeAddress> newRegions = new ArrayList<CellRangeAddress>();
    for (int i = sheet.getNumMergedRegions()-1; i >=0 ; i--) {
      CellRangeAddress r = sheet.getMergedRegion(i);
      if (r.getFirstRow() == forstart && r.getLastRow() == forend) {
        CellRangeAddress n_r = new CellRangeAddress(StartRowNo-1, EndRowNo-1, r.getFirstColumn(), r.getLastColumn());
        newRegions.add(n_r);
        sheet.removeMergedRegion(i);
        Cell fromCell = WorkbookUtils.getCell(sheet, forstart, r.getFirstColumn());
        Cell toCell = WorkbookUtils.getCell(sheet, StartRowNo-1, r.getFirstColumn());
        toCell.setCellStyle(fromCell.getCellStyle());
        if(fromCell.getCellType()==CellType.FORMULA){
          toCell.setCellFormula(fromCell.getCellFormula());
        }else{
          toCell.setCellType(fromCell.getCellType());
        }
        switch (fromCell.getCellType()) {
          case BOOLEAN:
            toCell.setCellValue(fromCell.getBooleanCellValue());
            break;
          case FORMULA:
            toCell.setCellFormula(fromCell.getCellFormula());
            break;
          case NUMERIC:
            toCell.setCellValue(fromCell.getNumericCellValue());
            break;
          case STRING:
            toCell.setCellValue(fromCell.getStringCellValue());
            break;
          default:
        }
      }
    }
    for (CellRangeAddress region :newRegions ) {
      sheet.addMergedRegion(region);
    }


    // delete #foreach #end block
    for (int rownum = forstart; rownum <= forend; rownum++) {
      sheet.removeRow(WorkbookUtils.getRow(rownum, sheet));
    }
    
    // remove merged region in forstart & forend    
    for (int i=0; i<sheet.getNumMergedRegions(); i++) {
        CellRangeAddress r = sheet.getMergedRegion(i);
    	if (r.getFirstRow()>=forstart && r.getLastRow()<=forend) {
    		sheet.removeMergedRegion(i);
    		// we have to back up now since we removed one
    		i = i - 1;
    	}
    }
    int startRow = forstart + 1;
    int endRow = sheet.getLastRowNum();
    if(startRow < endRow){
        sheet.shiftRows(forend + 1, sheet.getLastRowNum(), -(forend - forstart + 1), true, true);
    }
    return new int[] { ExcelParser.getSkipNum(forstart, forend),
        ExcelParser.getShiftNum(old_forend, forstart), 1 };
  }

  public String getTagName() {
    return KEY_FOREACH;
  }

  public boolean hasEndTag() {
    return true;
  }
}
