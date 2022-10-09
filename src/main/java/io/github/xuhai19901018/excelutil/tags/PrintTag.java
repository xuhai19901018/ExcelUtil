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

import io.github.xuhai19901018.excelutil.ExcelParser;
import io.github.xuhai19901018.excelutil.ExcelUtils;
import io.github.xuhai19901018.excelutil.WorkbookUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Iterator;
import java.util.StringTokenizer;


/***
 * 2022年10月9日
 * 打印区域设置
 * @author xuhai
 */
public class PrintTag implements ITag {

  public static final String KEY_PRINT = "#print";

  public int[] parseTag(Object context, Sheet sheet, Row curRow, Cell curCell) throws Exception {
    String expr = curCell.getStringCellValue();
    StringTokenizer st = new StringTokenizer(expr, " ");
    int with = 0;

    int pos = 0;
    while (st.hasMoreTokens()) {
      String str = st.nextToken();
      if (pos == 1) {
        with = Integer.parseInt(str);
      }
      pos++;
    }
    ExcelUtils.addValue(context, "printAreaEndRowNo", curRow.getRowNum()-1);
    ExcelUtils.addValue(context, "printAreaStartColNo", curCell.getColumnIndex());
    ExcelUtils.addValue(context, "printAreaColumns", with-1);

    sheet.removeRow(WorkbookUtils.getRow(curRow.getRowNum(), sheet));
    return new int[] { 1, -1, 1 };
  }

  public String getTagName() {
    return KEY_PRINT;
  }

  public boolean hasEndTag() {
    return false;
  }
}
