/*
 * Copyright 2003-2005 ExcelUtils http://excelutils.sourceforge.net
 * Created on 2005-7-5
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

import java.io.*;
import java.util.Map;

import javax.servlet.ServletContext;

import com.aspose.cells.IndividualFontConfigs;
import com.aspose.cells.License;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import org.apache.commons.beanutils.DynaBean;
import org.apache.commons.beanutils.LazyDynaBean;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


/**
 * <p>
 * <b>ExcelUtils</b> is a class which parse the excel report template
 * </p>
 *
 * @author rainsoft
 * @version $Revision: 1.11 $ $Date: 2005/11/02 10:31:22 $
 */
public class ExcelUtils {
    static ThreadLocal context = new ThreadLocal();

    /**
     * parse the excel template and output excel to outputStream.
     *
     * @param ctx     ServletContext
     * @param config  Excel Template Name
     * @param context All Data
     * @param out     OutputStream
     * @throws Exception
     */
    public static void export(ServletContext ctx, String config, Object context, OutputStream out) throws Exception {

        Workbook wb = WorkbookUtils.openWorkbook(ctx, config);
        parseWorkbook(context, wb);
        wb.write(out);

    }

    /**
     * parse the excel template in a sheet and output excel to outputStream.
     *
     * @param ctx        ServletContext
     * @param config     file name
     * @param sheetIndex sheetIndex
     * @param context    data object
     * @param out        OutputStream
     * @throws Exception
     */
    public static void export(ServletContext ctx, String config, int sheetIndex, Object context, OutputStream out)
            throws Exception {

        Workbook wb = WorkbookUtils.openWorkbook(ctx, config);
        parseWorkbook(context, wb, sheetIndex);
        wb.write(out);

    }

    /**
     * parse the excel template and output excel to outputStream in default
     * context.
     *
     * @param ctx    ServletContext
     * @param config file name
     * @param out    OutputStream
     * @throws Exception
     */
    public static void export(ServletContext ctx, String config, OutputStream out) throws Exception {

        export(ctx, config, getContext(), out);

    }

    /**
     * parse the excel template in a sheet and output excel to outputStream in
     * default context.
     *
     * @param ctx        ServletContext
     * @param config     file name
     * @param sheetIndex sheetIndex
     * @param out        OutputStream
     * @throws Exception
     */
    public static void export(ServletContext ctx, String config, int sheetIndex, OutputStream out) throws Exception {

        export(ctx, config, sheetIndex, getContext(), out);

    }

    /**
     * parse excel and export
     *
     * @param fileName file name
     * @param context  data object
     * @param out      OutputStream
     * @throws Exception
     */
    public static void export(String fileName, Object context, OutputStream out) throws Exception {

        Workbook wb = WorkbookUtils.openWorkbook(fileName);
        parseWorkbook(context, wb);
        wb.write(out);

    }

    /**
     * parse exel and export
     *
     * @param fileName   file name
     * @param sheetIndex sheetIndex
     * @param context    data object
     * @param out        OutputStream
     * @throws Exception
     */
    public static void export(String fileName, int sheetIndex, Object context, OutputStream out) throws Exception {

        Workbook wb = WorkbookUtils.openWorkbook(fileName);
        parseWorkbook(context, wb, sheetIndex);
        wb.write(out);

    }

    /**
     * parse excel and export excel
     *
     * @param fileName file name
     * @param out      OutputStream
     * @throws Exception
     */
    public static void export(String fileName, OutputStream out) throws Exception {
        export(fileName, getContext(), out);

    }

    public static void exportPdf(String fileName, OutputStream out) throws Exception {
        String tempFile = System.getProperty("java.io.tmpdir") + File.separator +new File(fileName).getName() + ".xlsx";

        export(fileName, getContext(), new FileOutputStream(new File(tempFile)));
//      export(fileName, getContext(), new FileOutputStream(new File("D:\\home\\t4.xlsm")));
        InputStream is = ExcelUtils.class.getClassLoader().getResourceAsStream("pdfLicense/license.xml");//这个文件应该是类似于密码验证(证书？)，用于获得去除水印的权限
        License aposeLic = new License();
        aposeLic.setLicense(is);
        com.aspose.cells.Workbook wbk = new com.aspose.cells.Workbook(tempFile);
        wbk.save(out, SaveFormat.PDF);


    }

    public static void exportPdf(String fileName, OutputStream out, String fontFolder) throws Exception {

        com.aspose.cells.FontConfigs.setFontFolder(fontFolder, true);

        exportPdf(fileName, out);

    }


    /**
     * parse excel and export excel
     *
     * @param fileName   file name
     * @param sheetIndex sheetIndex
     * @param out        OutputStream
     * @throws Exception
     */
    public static void export(String fileName, int sheetIndex, OutputStream out) throws Exception {
        export(fileName, sheetIndex, getContext(), out);

    }

    /**
     * @param inputStream file input stream
     * @param context     data object
     * @param out         OutputStream
     * @throws Exception
     */
    public static void export(InputStream inputStream, Object context, OutputStream out) throws Exception {
        Workbook wb = WorkbookUtils.openWorkbook(inputStream);
        parseWorkbook(context, wb);
        wb.write(out);

    }

    /**
     * parse workbook
     *
     * @param context data object
     * @param wb      Workbook
     * @throws Exception
     */
    public static void parseWorkbook(Object context, Workbook wb) throws Exception {

        int sheetCount = wb.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            Sheet sheet = wb.getSheetAt(sheetIndex);
            parseSheet(context, sheet);
            // sheet名可作为表达式
            String sheetName = sheet.getSheetName();
            sheet.getWorkbook().setSheetName(sheetIndex, (String) ExcelParser.parseStr(context, sheetName));

            sheet.setForceFormulaRecalculation(true);
            try {//尝试设置打印区域
                wb.setPrintArea(sheetIndex, (Integer) ExcelParser.getValue(context, "printAreaStartColNo"), (Integer) ExcelParser.getValue(context, "printAreaStartColNo") + (Integer) ExcelParser.getValue(context, "printAreaColumns"), 0, (Integer) ExcelParser.getValue(context, "printAreaEndRowNo"));
                sheet.setRepeatingRows(CellRangeAddress.valueOf((String) ExcelParser.getValue(context, "repeatingRows")));
            } catch (Exception e) {
            }
        }

    }

    /**
     * parse Workbook
     *
     * @param context    data object
     * @param wb         Workbook
     * @param sheetIndex sheetIndex
     * @throws Exception
     */
    public static void parseWorkbook(Object context, Workbook wb, int sheetIndex) throws Exception {

        Sheet sheet = wb.getSheetAt(sheetIndex);
        if (null != sheet) {
            parseSheet(context, sheet);
            // sheet名可作为表达式
            String sheetName = sheet.getSheetName();
            sheet.getWorkbook().setSheetName(sheetIndex, (String) ExcelParser.parseStr(context, sheetName));
            sheet.setForceFormulaRecalculation(true);
            try {//尝试设置打印区域
                wb.setPrintArea(sheetIndex, (Integer) ExcelParser.getValue(context, "printAreaStartColNo"), (Integer) ExcelParser.getValue(context, "printAreaStartColNo") + (Integer) ExcelParser.getValue(context, "printAreaColumns"), 0, (Integer) ExcelParser.getValue(context, "printAreaEndRowNo"));
            } catch (Exception e) {
            }
        }

        int i = 0;
        while (i++ < sheetIndex) {
            wb.removeSheetAt(0);
        }

        i = 1;
        while (i < wb.getNumberOfSheets()) {
            wb.removeSheetAt(i);
        }

    }

    /**
     * parse Excel Template File
     *
     * @param context datasource
     * @param sheet   Workbook sheet
     */
    public static void parseSheet(Object context, Sheet sheet) throws Exception {
        try {
            ExcelParser.parse(context, sheet, sheet.getFirstRowNum(), sheet.getLastRowNum());
            ExcelParser.parseChart(context, sheet);
        } finally {
            ExcelUtils.context.set(null);
        }
    }

    public static void addService(Object context, String key, Object service) {
        addValue(context, key, service);
    }

    public static void addService(String key, Object service) {
        addValue(key, service);
    }

    /**
     * add a object to context
     *
     * @param context must be a DynaBean or Map type
     * @param value   data
     */
    public static void addValue(Object context, String key, Object value) {
        if (context instanceof DynaBean) {
            ((DynaBean) context).set(key, value);
        } else if (context instanceof Map) {
            ((Map) context).put(key, value);
        }
    }

    /**
     * add a object to default context
     *
     * @param key   key
     * @param value value
     */
    public static void addValue(String key, Object value) {
        getContext().set(key, value);
    }

    /**
     * register extended tag package, default is net.sf.excelutils.tags
     *
     * @param packageName package name
     */
    public synchronized static void registerTagPackage(String packageName) {
        ExcelParser.tagPackageMap.put(packageName, packageName);
    }

    /**
     * get a global context, it's thread safe
     *
     * @return DynaBean
     */
    public static DynaBean getContext() {
        DynaBean ctx = (DynaBean) context.get();
        if (null == ctx) {
            ctx = new LazyDynaBean();
            setContext(ctx);
        }
        return ctx;
    }

    /**
     * set global context
     *
     * @param ctx DynaBean
     */
    public static void setContext(DynaBean ctx) {
        context.set(ctx);
    }

    /**
     * can value be show
     *
     * @param value data
     * @return boolean
     */
    public static boolean isCanShowType(Object value) {
        if (null == value) return false;
        String valueType = value.getClass().getName();
        return "java.lang.String".equals(valueType) || "java.lang.Double".equals(valueType)
                || "java.lang.Integer".equals(valueType) || "java.lang.Boolean".equals(valueType)
                || "java.sql.Timestamp".equals(valueType) || "java.util.Date".equals(valueType)
                || "java.lang.Byte".equals(valueType) || "java.math.BigDecimal".equals(valueType)
                || "java.math.BigInteger".equals(valueType) || "java.lang.Float".equals(valueType)
                || value.getClass().isPrimitive();
    }
}