/*
 * Copyright 2003-2005 ExcelUtils http://excelutils.sourceforge.net
 * Created on 2005-6-18
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
package io.github.xuhai19901018.excelutil;

import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import org.apache.commons.beanutils.DynaClass;
import org.apache.commons.beanutils.DynaProperty;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import io.github.xuhai19901018.excelutil.tags.ITag;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;


/**
 * <p>
 * <b>ExcelParser </b> is a class which can parse the Excel template <BR>
 * to genrate reports.
 * </p>
 *
 * @author rainsoft
 * @version $Revision: 1.7 $ $Date: 2005/11/03 03:35:19 $
 */
public class ExcelParser {

    public static final String VALUED_DELIM = "${";

    public static final String VALUED_DELIM2 = "}";

    public static final String KEY_TAG = "#";

    public static Map tagPackageMap = new HashMap();

    private static Map tagMap = new HashMap();

    static {
        tagPackageMap.put(ITag.class.getPackage().getName(), ITag.class.getPackage().getName());
    }

    /***
     * 图表填充
     * @param context
     * @param sheet
     * @return
     * @throws Exception
     */
    public static void parseChart(Object context, Sheet sheet) {

        if (sheet instanceof XSSFSheet){
            XSSFSheet xssfSheet = (XSSFSheet)sheet;
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            if (null !=drawing){
                List<XSSFChart> charts = drawing.getCharts();
                if (null != charts && charts.size()>0)
                    for (XSSFChart chart : charts) {
                        if (null != chart.getTitleText()){
                            String expr = chart.getTitleText().toString();
                            if (null!= expr)
                                chart.setTitleText((String) parseStr(context, expr));
                        }
                    }
            }
        }
    }

    /**
     * parse the Excel template
     *
     * @param context data object
     * @param sheet   Excel sheet
     * @param fromRow the start
     * @param toRow   the end
     * @return int skip number
     * @throws Exception
     */
    public static int parse(Object context, Sheet sheet, int fromRow, int toRow) throws Exception {
        int[] shift = new int[]{0, 0, 0}; // {SkipNum, ShiftNum, break flag}
        int shiftCount = 0;

        int rownum = fromRow;
        while (rownum <= toRow) {
            // shift
            rownum += shift[1] + shift[0];
            toRow += shift[1];
            if (rownum > toRow)
                break;

            shift[0] = 0;
            shift[1] = 0;
            shift[2] = 0;
            Row row = sheet.getRow(rownum);
            // set current row number
            ExcelUtils.addValue(context, "currentRowNo", new Integer(rownum + 1));
            if (null == row) {
                rownum++;
                continue;
            }

            for (short colnum = row.getFirstCellNum(); colnum <= row.getLastCellNum(); colnum++) {
                if (colnum < 0) {
                    continue;
                }
                Cell cell = row.getCell(colnum);
                if (null == cell) {
                    continue;
                }
                if (cell.getCellType() != CellType.STRING) {
                    continue;
                }
                // if the cell is null then continue
                String cellstr = cell.getStringCellValue();
                if (null == cellstr || "".equals(cellstr)) {
                    continue;
                }

                ITag tag = getTagClass(cellstr);
                if (null != tag) {
                    shift = tag.parseTag(context, sheet, row, cell);
                } else {
                    parseCell(context, sheet, row, cell);
                }

                shiftCount += shift[1];
                if (shift[2] == 1)
                    break;
            }
            rownum++;
        }
        return shiftCount;
    }

    /**
     * get a instance by the tag name.
     *
     * @param str tag name
     * @return ITag instance
     */
    public static ITag getTagClass(String str) {
        String tagName = "";
        int keytag = str.indexOf(KEY_TAG);
        if (keytag < 0)
            return null;
        if (!(keytag < str.length() - 1))
            return null;
        String tagRight = str.substring(keytag + 1, str.length());
        if (tagRight.startsWith(KEY_TAG))
            return null;

        str = str.substring(str.indexOf(KEY_TAG) + KEY_TAG.length(), str.length());
        StringTokenizer st = new StringTokenizer(str, " ");
        if (st.hasMoreTokens()) {
            tagName = st.nextToken();
        }
        tagName = tagName.substring(0, 1).toUpperCase() + tagName.substring(1, tagName.length());
        tagName += "Tag";

        // find in tagMap first, if not exist, search in the package
        ITag tag = (ITag) tagMap.get(tagName);
        if (tag == null) {
            Iterator tagPackages = tagPackageMap.values().iterator();
            // seach the tag class
            while (tagPackages.hasNext() && null == tag) {
                String packageName = (String) tagPackages.next();
                try {
                    Class clazz = Class.forName(packageName + "." + tagName);
                    tag = (ITag) clazz.newInstance();
                } catch (Exception e) {
                    tag = null;
                }
            }
            if (tag != null) {
                // find it, cache it
                tagMap.put(tagName, tag);
            }
        }
        return tag;
    }

    /**
     * get the value from context by the expression
     *
     * @param expr    expression
     * @param context data object
     * @return Object the value of the expr
     */
    public static Object getValue(Object context, String expr) {
        Object value = null;
        try {
            value = PropertyUtils.getProperty(context, expr);
        } catch (Exception e) {
            return null;
        }
        return value;
    }

    /**
     * parse the expression ${model[${index}]} only one ${} and startWith ${
     * endWith }
     *
     * @param context data object
     * @param expr    expression
     * @return the value of the expression
     */
    public static Object parseExpr(Object context, String expr) {

        int indexValued = expr.indexOf(VALUED_DELIM);
        int indexValued2 = expr.lastIndexOf(VALUED_DELIM2);

        Object value = null;

        if (indexValued == 0 && indexValued2 > 0) {
            String property = expr.substring(indexValued + VALUED_DELIM.length(), indexValued2);
            if (property.startsWith("!")) {
                property = property.substring(1, property.length());
            }

            if (property.indexOf(VALUED_DELIM) >= 0) {
                Object pro = parseStr(context, property);
                if (null != pro)
                    property = pro.toString();
            }
            // get the value by expression
            value = getValue(context, property);
        } else if (indexValued > 0 && indexValued2 > 0) {
            value = parseStr(context, expr);
        }

        return value;
    }

    /**
     * parse complex expression ${${}}aaa${}
     *
     * @param context data object
     * @param str     expression
     * @return value of the expression
     */
    public static Object parseStr(Object context, String str) {
        return parseStr(context, str, false);
    }

    /**
     * parse complex expression ${${}}aaa${}
     *
     * @param context data object
     * @param str     expression
     * @param quot    string needs quotation or not
     * @return value of the expression
     */
    public static Object parseStr(Object context, String str, boolean quot) {

        int exprCount = 0;
        int valueFrom = -1;
        int valueTo = -1;
        int valueCount = 0;
        int pos = 0;
        Object value = null;

        int indexValued = str.indexOf(VALUED_DELIM);
        int indexValued2 = str.lastIndexOf(VALUED_DELIM2);

        boolean bJustExpr = str.length() == indexValued2 + VALUED_DELIM2.length() - indexValued;

        while (pos < str.length()) {
            if (pos + VALUED_DELIM.length() <= str.length()) {
                if (VALUED_DELIM.equals(str.substring(pos, pos + VALUED_DELIM.length()))) {
                    if (valueCount == 0) {
                        valueFrom = pos;
                    }
                    valueCount++;
                    pos = pos + VALUED_DELIM.length();
                    continue;
                }
            }

            if (VALUED_DELIM2.equals(str.substring(pos, pos + VALUED_DELIM2.length()))) {
                valueCount--;
                if (valueCount == 0) {
                    valueTo = pos;
                    String expr = str.substring(valueFrom, valueTo + VALUED_DELIM2.length());
                    value = parseExpr(context, expr);
                    exprCount++;
                    // replace the string
                    StringBuffer sbuf = new StringBuffer(str);
                    if (null != value) {
                        String rep = value.toString();
                        // need quotation
                        if (quot) {
                            rep = "\"" + rep + "\"";
                        }
                        sbuf.replace(valueFrom, valueTo + VALUED_DELIM2.length(), rep);
                        pos += VALUED_DELIM2.length() + value.toString().length() - expr.length();
                    } else {
                        String rep = "";
                        // need quotation
                        if (quot) {
                            rep = "\"" + rep + "\"";
                        }
                        sbuf.replace(valueFrom, valueTo + VALUED_DELIM2.length(), rep);
                        pos += VALUED_DELIM2.length() + 0 - expr.length();
                    }
                    str = sbuf.toString();
                    continue;
                } else {
                    pos += VALUED_DELIM2.length();
                    continue;
                }
            }
            pos++;
        }

        if (exprCount == 1 && bJustExpr) {
            if (null != value) {
                if (quot && "java.lang.String".equals(value.getClass().getName())) {
                    return "\"" + value.toString() + "\"";
                }
                return value;
            }
            return value;
        } else {
            return str;
        }
    }

    /**
     * parse the cell
     *
     * @param context data object
     * @param cell    excel cell
     */
    public static void parseCell(Object context, Sheet sheet, Row row, Cell cell) {

        String str = cell.getStringCellValue();
        if (null == str || "".equals(str)) {
            return;
        }

        if (str.indexOf(VALUED_DELIM) < 0)
            return;

        boolean bJustExpr = str.length() == (str.length() - str.lastIndexOf(VALUED_DELIM));
        boolean bMerge = "!".equals(str.substring(str.indexOf(VALUED_DELIM) + VALUED_DELIM.length(), str
                .indexOf(VALUED_DELIM)
                + VALUED_DELIM.length() + 1));

        if (str.indexOf(VALUED_DELIM) < 0)
            return;

        Object value = parseStr(context, str);

        // replace the cell
        if (null != value) {
            String className = value.getClass().getName();

            if (bJustExpr && ("java.lang.String".equals(className))) {
                cell.setCellValue(String.valueOf(value));
                cell.setCellType(CellType.STRING);
            } else if (bJustExpr && String.valueOf(value).matches("-?[0-9]+.?[0-9]*")) {
                cell.setCellValue(Double.parseDouble(String.valueOf(value)));
                cell.setCellType(CellType.NUMERIC);
            } else if (bJustExpr && "java.util.Date".equals(className)) {
                cell.setCellValue((Date) value);
                cell.setCellType(CellType.NUMERIC);
            } else if (bJustExpr && "java.lang.Boolean".equals(className)) {
                cell.setCellValue(((Boolean) value).booleanValue());
                cell.setCellType(CellType.BOOLEAN);
            } else {
                cell.setCellValue(String.valueOf(value));
                cell.setCellType(CellType.STRING);
            }
        } else {
            cell.setCellValue("");
            cell.setCellType(CellType.BLANK);
        }

        // merge the cell that has a "!" character at the expression
        if (row.getRowNum() - 1 >= sheet.getFirstRowNum() && bMerge) {
            Row lastRow = WorkbookUtils.getRow(row.getRowNum() - 1, sheet);
            Cell lastCell = WorkbookUtils.getCell(lastRow, cell.getColumnIndex());
            boolean canMerge = false;
            if (lastCell.getCellType() == cell.getCellType()) {
                switch (cell.getCellType()) {
                    case STRING:
                        canMerge = lastCell.getStringCellValue().equals(cell.getStringCellValue());
                        break;
                    case BOOLEAN:
                        canMerge = lastCell.getBooleanCellValue() == cell.getBooleanCellValue();
                        break;
                    case NUMERIC:
                        canMerge = lastCell.getNumericCellValue() == cell.getNumericCellValue();
                        break;
                }
            }
            if (canMerge) {
                sheet.addMergedRegion(new CellRangeAddress(lastRow.getRowNum(), lastCell.getColumnIndex(), row.getRowNum(), cell.getColumnIndex()));
            }
        }

    }

    /**
     * get properties of the JavaBean
     *
     * @param clazz JavaBean
     * @return fields in the javabean
     */
    public static Field[] getBeanProperties(Class clazz) {
        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getMethods();
        String m = "";

        for (int i = 0; i < methods.length; i++) {
            m += methods[i].getName() + ",";
        }

        List flist = new ArrayList();
        for (int i = 0; i < fields.length; i++) {
            if (m.indexOf("get" + fields[i].getName().substring(0, 1).toUpperCase()
                    + fields[i].getName().substring(1, fields[i].getName().length())) >= 0) {
                flist.add(fields[i]);
            }
        }
        Field[] result = new Field[flist.size()];
        flist.toArray(result);
        return result;
    }

    public static DynaProperty[] getBeanProperties(DynaClass clazz) {
        DynaProperty dynaProperties[] = clazz.getDynaProperties();
        return dynaProperties;
    }

    /**
     * get Iterator from the object
     *
     * @param collection collection
     * @return Iterator of the object
     */
    public static Iterator getIterator(Object collection) {
        Iterator iterator = null;
        if (collection.getClass().isArray()) {
            try {
                // If we're lucky, it is an array of objects
                // that we can iterate over with no copying
                iterator = Arrays.asList((Object[]) collection).iterator();
            } catch (ClassCastException e) {
                // Rats -- it is an array of primitives
                int length = Array.getLength(collection);
                ArrayList c = new ArrayList(length);
                for (int i = 0; i < length; i++) {
                    c.add(Array.get(collection, i));
                }
                iterator = c.iterator();
            }
        } else if (collection instanceof Collection) {
            iterator = ((Collection) collection).iterator();
        } else if (collection instanceof Iterator) {
            iterator = (Iterator) collection;
        } else if (collection instanceof Map) {
            iterator = ((Map) collection).entrySet().iterator();
        }
        return iterator;
    }

    /*
     * (non-Javadoc)
     *
     * @see net.sf.excelutils.webwork.ExcelLoader#getExcel(java.lang.String)
     */
    public static int getSkipNum(int tagstart, int tagend) {
        return tagend - tagstart;
    }

    /*
     * (non-Javadoc)
     *
     * @see net.sf.excelutils.webwork.ExcelLoader#getExcel(java.lang.String)
     */
    public static int getShiftNum(int old_tagend, int tagstart) {
        return tagstart - old_tagend - 1;
    }
}
