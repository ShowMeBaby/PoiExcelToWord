package com.ExcelToWord.utils;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

import com.ExcelToWord.entity.CellEntity;
import com.ExcelToWord.entity.VarMapEntity;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel操作工具类
 *
 * @author ChaiXY
 */
public class ExcelUtils {
    public static final String OFFICE_EXCEL_XLS = "xls";
    public static final String OFFICE_EXCEL_XLSX = "xlsx";

    /**
     * 读取指定Sheet也的内容
     *
     * @param filepath filepath 文件全路径
     * @param sheetNo  sheet序号,从0开始,如果读取全文sheetNo设置null
     */
    public static String readExcel(String filepath, Integer sheetNo)
            throws EncryptedDocumentException, InvalidFormatException, IOException {
        StringBuilder sb = new StringBuilder();
        Workbook workbook = getWorkbook(filepath);
        if (workbook != null) {
            if (sheetNo == null) {
                int numberOfSheets = workbook.getNumberOfSheets();
                for (int i = 0; i < numberOfSheets; i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    if (sheet == null) {
                        continue;
                    }
                    sb.append(readExcelSheet(sheet));
                }
            } else {
                Sheet sheet = workbook.getSheetAt(sheetNo);
                if (sheet != null) {
                    sb.append(readExcelSheet(sheet));
                }
            }
        }
        return sb.toString();
    }

    /**
     * 将变量表转换为Map
     *
     * @param workbook 变量表工作簿
     * @return 变量表的KV值
     */
    public static Map<String, VarMapEntity> getSheetVarMap(Workbook workbook) throws Exception {
        Sheet sheet = workbook.getSheet("变量表");
        Map<String, VarMapEntity> map = new HashMap<>();
        if (sheet != null) {
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {// 遍历行
                Row row = sheet.getRow(i);
                String rowTitle = getRowTitle(row);
                if (!rowTitle.isEmpty()) {
                    String cellValue = getCellValue(row.getCell(1));
                    String[] cellValueArr = cellValue.split(" ");
                    if (cellValueArr.length == 2 && cellValueArr[1].split(":").length == 2) {
                        // 自定义指令引用表格
                        ArrayList<ArrayList<CellEntity>> colRowToTable = getColRowToTable(workbook.getSheet(cellValueArr[0]), getTablePoint(cellValueArr[1]));
                        VarMapEntity varMapEntity = new VarMapEntity("Table");
                        varMapEntity.setTableValue(colRowToTable);
                        map.put("${" + rowTitle + "}", varMapEntity);
                    } else {
                        VarMapEntity varMapEntity = new VarMapEntity("String");
                        varMapEntity.setStringValue(cellValue);
                        map.put("${" + rowTitle + "}", varMapEntity);
                    }
                }
            }
        } else {
            throw new Exception("变量表不存在");
        }
        return map;
    }

    /**
     * 根据坐标获取excel中的表格
     *
     * @param tablePoint 变量表工作簿
     * @return 变量表的KV值
     */
    private static ArrayList<ArrayList<CellEntity>> getColRowToTable(Sheet sheet, CellEntity tablePoint) {
        // 当前表所有合并单元坐标集合
        List<CellEntity> sheetMergedRegions = getSheetMergedRegions(sheet);
        // 表格的所有行数据
        ArrayList<ArrayList<CellEntity>> allRowList = new ArrayList<>();
        // 获取每行的值
        for (int i = tablePoint.getStartRowIndex(); i < tablePoint.getEndRowIndex(); i++) {
            Row row = sheet.getRow(i);
            ArrayList<CellEntity> rowList = new ArrayList<>();

            for (int j = tablePoint.getStartColumnIndex(); j < tablePoint.getEndColumnIndex(); j++) {
                String cellValue = getCellValue(row.getCell(j));
                CellEntity cellEntity = new CellEntity(cellValue, i, i, j, j);
                // 查询当前单元格是否是合并单元格
                List<CellEntity> collects = sheetMergedRegions.stream().filter(o -> cellEntity.getStartRowIndex() >= o.getStartRowIndex() && cellEntity.getEndRowIndex() <= o.getEndRowIndex() && cellEntity.getStartColumnIndex() >= o.getStartColumnIndex() && cellEntity.getEndColumnIndex() <= o.getEndColumnIndex()).collect(Collectors.toList());
                if (collects.size() >= 1) {
                    // 该单元格是合并单元格
                    CellEntity collect = collects.get(0);
                    cellEntity.setMergedRegions(new int[]{collect.getStartRowIndex(), collect.getEndRowIndex(), collect.getStartColumnIndex(), collect.getEndColumnIndex()});
                    if (!cellEntity.getValue().isEmpty()){
                        // 如果值不为空,则将该合并单元格标记为已填充
                        collect.setToMergedRegion(true);
                    }
                }
                rowList.add(cellEntity);
            }
            List<String> rowListHasValue = rowList.stream().filter(o -> !o.getValue().isEmpty()).map(CellEntity::getValue).collect(Collectors.toList());
            // 整行为空,则代表整行是上下某行的合并单元格,我们不需要直接过滤掉
            if (rowListHasValue.size() > 0) {
                for (CellEntity cellEntity : rowList) {
                    List<CellEntity> collects = sheetMergedRegions.stream().filter(o -> cellEntity.getStartRowIndex() >= o.getStartRowIndex() && cellEntity.getEndRowIndex() <= o.getEndRowIndex() && cellEntity.getStartColumnIndex() >= o.getStartColumnIndex() && cellEntity.getEndColumnIndex() <= o.getEndColumnIndex()).collect(Collectors.toList());
                    // 如果某一格没值,同时又是合并单元格,则获取合并单元格值补全
                    if (cellEntity.getValue().isEmpty() && cellEntity.getMergedRegions().length == 4 && !collects.get(0).isToMergedRegion()) {
                        // 该合并单元格的值已经填充过,则不在填充
                        collects.get(0).setToMergedRegion(true);
                        cellEntity.setValue(getMergedRegionValue(sheet, cellEntity.getStartRowIndex(), cellEntity.getStartColumnIndex()));
                    }
                }
                allRowList.add(rowList);
            }
        }
        return allRowList;
    }

    /**
     * 获取sheet所有合并单元格的坐标
     *
     * @param sheet 工作表
     * @return 变量表的KV值
     */
    public static List<CellEntity> getSheetMergedRegions(Sheet sheet) {
        List<CellEntity> list = new ArrayList<>();
        //获得一个 sheet 中合并单元格的数量
        int sheetmergerCount = sheet.getNumMergedRegions();
        //遍历所有的合并单元格
        for (int i = 0; i < sheetmergerCount; i++) {
            //获得合并单元格保存进list中
            CellRangeAddress ca = sheet.getMergedRegion(i);

            list.add(getTablePoint(ca.formatAsString()));
        }
        return list;
    }

    /**
     * 获取合并单元格的值
     */
    public static String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }
        return "";
    }

    /**
     * 把A1:F5这种的坐标转换为POI所用的列数及行数
     *
     * @param tablePoint 表格坐标
     * @return 开始行数, 开始列数, 结束行数, 结束列数
     */
    private static CellEntity getTablePoint(String tablePoint) {
        String[] rowColumnPoint = tablePoint.split(":");
        // 开始坐标
        String start_str_string = Arrays.stream(rowColumnPoint[0].split("\\d")).filter(o -> !o.isEmpty()).collect(Collectors.toList()).get(0);//  \d 为正则表达式表示[0-9]数字
        String start_num_string = Arrays.stream(rowColumnPoint[0].split("\\D")).filter(o -> !o.isEmpty()).collect(Collectors.toList()).get(0); // \D 为正则表达式表示非数字
        int start_columnIndex = getExcelColumnIndex(start_str_string);
        int start_rowIndex = getExcelRowIndex(start_num_string);
        // 结束坐标
        String end_str_string = Arrays.stream(rowColumnPoint[1].split("\\d")).filter(o -> !o.isEmpty()).collect(Collectors.toList()).get(0);
        String end_num_string = Arrays.stream(rowColumnPoint[1].split("\\D")).filter(o -> !o.isEmpty()).collect(Collectors.toList()).get(0);
        int end_columnIndex = getExcelColumnIndex(end_str_string);
        // 友好型修正,结束行默认+1,使表格符号更符合直觉,如: A1:H10 包括第10行,而不是到第10行结束
        int end_rowIndex = getExcelRowIndex(end_num_string) + 1;
        // 修正写反的坐标
        if (start_columnIndex > end_columnIndex) {
            int temp_columnIndex = start_columnIndex;
            start_columnIndex = end_columnIndex;
            end_columnIndex = temp_columnIndex;
        }
        if (start_rowIndex > end_rowIndex) {
            int temp_rowIndex = start_rowIndex;
            start_rowIndex = end_rowIndex;
            end_rowIndex = temp_rowIndex;
        }
        return new CellEntity("", start_rowIndex, end_rowIndex, start_columnIndex, end_columnIndex);
    }

    /**
     * 获取列数字母对应的列数
     *
     * @param str_string 列数字母字符串
     */
    private static int getExcelColumnIndex(String str_string) {
        str_string = str_string.toUpperCase();
        int columnIndex = 0;
        char[] letters = str_string.toCharArray();
        for (char let : letters) {
            columnIndex = (int) let - 64 + columnIndex * 26;
        }
        return columnIndex;
    }

    /**
     * 获取行数字符串对应的行数
     *
     * @param num_string 行数字符串
     */
    private static int getExcelRowIndex(String num_string) {
        return Integer.parseInt(num_string) - 1;
    }

    /**
     * 根据文件路径获取Workbook对象
     *
     * @param filepath 文件全路径
     */
    public static Workbook getWorkbook(String filepath)
            throws EncryptedDocumentException, IOException {
        InputStream is = null;
        Workbook wb = null;

        if (filepath.isEmpty()) {
            throw new IllegalArgumentException("文件路径不能为空");
        } else {
            String suffiex = getSuffiex(filepath);
            if (suffiex.isEmpty()) {
                throw new IllegalArgumentException("文件后缀不能为空");
            }
            if (OFFICE_EXCEL_XLS.equals(suffiex) || OFFICE_EXCEL_XLSX.equals(suffiex)) {
                try {
                    is = new FileInputStream(filepath);
                    wb = WorkbookFactory.create(is);
                } finally {
                    if (is != null) {
                        is.close();
                    }
                    if (wb != null) {
                        wb.close();
                    }
                }
            } else {
                throw new IllegalArgumentException("该文件非Excel文件");
            }
        }
        return wb;
    }

    /**
     * 获取后缀
     *
     * @param filepath filepath 文件全路径
     */
    private static String getSuffiex(String filepath) {
        if (filepath.isEmpty()) {
            return "";
        }
        int index = filepath.lastIndexOf(".");
        if (index == -1) {
            return "";
        }
        return filepath.substring(index + 1);
    }

    private static String readExcelSheet(Sheet sheet) {
        StringBuilder sb = new StringBuilder();
        if (sheet != null) {
            int rowNos = sheet.getLastRowNum();// 得到excel的总记录条数
            for (int i = 0; i <= rowNos; i++) {// 遍历行
                Row row = sheet.getRow(i);
                if (row != null) {
                    int columNos = row.getLastCellNum();// 表头总共的列数
                    for (int j = 0; j < columNos; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            sb.append(cell.getStringCellValue()).append(" ");
                        }
                    }
                }
            }
        }
        return sb.toString();
    }

    /**
     * 读取每行A列的值为变量名称
     *
     * @param row 行
     */
    public static String getRowTitle(Row row) {
        return getCellValue(row.getCell(0));
    }

    /**
     * 读取每行A列的值为变量名称
     *
     * @param cell 格
     * @return 格子的值
     */
    public static String getCellValue(Cell cell) {
        String result = "";
        /* 如果有百分号和时间没有转换出来,打印dataFormat然后把dataFormat丢进对应数组即可 */
        // 时间格式代码数组
        int[] dateDataFormatArr = {31};
        // 百分比格式代码数组
        int[] percentDataFormatArr = {9, 10};
        if (cell != null) {
            int dataFormat = cell.getCellStyle().getDataFormat();
            switch (cell.getCellType()) {
                case STRING:
                    // 按普通字符串算
                    result = cell.getStringCellValue().trim();
                    break;
                case NUMERIC:
                case FORMULA:
                    try {
                        double numericCellValue = cell.getNumericCellValue();
                        if (numericCellValue < 1 && numericCellValue > 0 && Arrays.binarySearch(percentDataFormatArr, dataFormat) >= 0) {
                            // 如果值小于1大于0,并且格式为9(百分百格式),则转为百分比
                            result = doubleToIntString(numericCellValue * 100) + "%";
                        } else if (DateUtil.isCellDateFormatted(cell) || Arrays.binarySearch(dateDataFormatArr, dataFormat) >= 0) {
                            // 如果是时间格式则转为时间
                            result = new SimpleDateFormat("yyyy年MM月dd日").format(cell.getDateCellValue());
                        } else {
                            // 按普通数字算
                            result = doubleToIntString(numericCellValue);
                        }
                    } catch (IllegalStateException e) {
                        // 公式只是单纯引用,结果为字符串,按字符串算
                        cell.setCellType(CellType.STRING);
                        result = cell.getStringCellValue().trim();
                    }
                    break;
                case BLANK:
                    // 没取到值,默认为空
                    cell.setCellType(CellType.STRING);
                    result = cell.getStringCellValue().trim();
                    break;
                case BOOLEAN:
                    // 布尔值直接转为字符串
                    result = String.valueOf(cell.getBooleanCellValue()).trim();
                    break;
                case _NONE:
                    System.out.println("发现空字符");
                case ERROR:
                    System.out.println("发现非法字符");
                default:
                    System.out.println("发现未处理字符");
            }
        }
        return result;
    }

    /**
     * 将double转换为string并去掉整数后的0
     *
     * @param dbl filepath 文件全路径
     */
    private static String doubleToIntString(double dbl) {
        if (dbl == (double) (int) dbl) {
            return (int) dbl + "";
        } else {
            return dbl + "";
        }
    }

    /**
     * 读取指定Sheet页的表头
     *
     * @param filepath filepath 文件全路径
     * @param sheetNo  sheet序号,从0开始,必填
     */
    public static Row readTitle(String filepath, int sheetNo)
            throws IOException, EncryptedDocumentException {
        Row returnRow = null;
        Workbook workbook = getWorkbook(filepath);
        if (workbook != null) {
            Sheet sheet = workbook.getSheetAt(sheetNo);
            returnRow = readTitle(sheet);
        }
        return returnRow;
    }

    /**
     * 读取指定Sheet页的表头
     */
    public static Row readTitle(Sheet sheet) {
        Row returnRow = null;
        int totalRow = sheet.getLastRowNum();// 得到excel的总记录条数
        for (int i = 0; i < totalRow; i++) {// 遍历行
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            returnRow = sheet.getRow(0);
            break;
        }
        return returnRow;
    }

    /**
     * 创建Excel文件
     *
     * @param filepath  filepath 文件全路径
     * @param sheetName 新Sheet页的名字
     * @param titles    表头
     * @param values    每行的单元格
     */
    public static boolean writeExcel(String filepath, String sheetName, List<String> titles,
                                     List<Map<String, Object>> values) {
        boolean success = false;
        OutputStream outputStream = null;
        if (filepath.isEmpty()) {
            throw new IllegalArgumentException("文件路径不能为空");
        } else {
            String suffiex = getSuffiex(filepath);
            if (suffiex.isEmpty()) {
                throw new IllegalArgumentException("文件后缀不能为空");
            }
            Workbook workbook;
            if ("xls".equals(suffiex.toLowerCase())) {
                workbook = new HSSFWorkbook();
            } else {
                workbook = new XSSFWorkbook();
            }
            // 生成一个表格
            Sheet sheet;
            if (sheetName.isEmpty()) {
                // name 为空则使用默认值
                sheet = workbook.createSheet();
            } else {
                sheet = workbook.createSheet(sheetName);
            }
            // 设置表格默认列宽度为15个字节
            sheet.setDefaultColumnWidth((short) 15);
            // 生成样式
            Map<String, CellStyle> styles = createStyles(workbook);
            // 创建标题行
            Row row = sheet.createRow(0);
            // 存储标题在Excel文件中的序号
            Map<String, Integer> titleOrder = new HashMap<>();
            for (int i = 0; i < titles.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(styles.get("header"));
                String title = titles.get(i);
                cell.setCellValue(title);
                titleOrder.put(title, i);
            }
            // 写入正文
            Iterator<Map<String, Object>> iterator = values.iterator();
            // 行号
            int index = 1;
            while (iterator.hasNext()) {
                row = sheet.createRow(index);
                Map<String, Object> value = iterator.next();
                for (Map.Entry<String, Object> map : value.entrySet()) {
                    // 获取列名
                    String title = map.getKey();
                    // 根据列名获取序号
                    int i = titleOrder.get(title);
                    // 在指定序号处创建cell
                    Cell cell = row.createCell(i);
                    // 设置cell的样式
                    if (index % 2 == 1) {
                        cell.setCellStyle(styles.get("cellA"));
                    } else {
                        cell.setCellStyle(styles.get("cellB"));
                    }
                    // 获取列的值
                    Object object = map.getValue();
                    // 判断object的类型
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    if (object instanceof Double) {
                        cell.setCellValue((Double) object);
                    } else if (object instanceof Date) {
                        String time = simpleDateFormat.format((Date) object);
                        cell.setCellValue(time);
                    } else if (object instanceof Calendar) {
                        Calendar calendar = (Calendar) object;
                        String time = simpleDateFormat.format(calendar.getTime());
                        cell.setCellValue(time);
                    } else if (object instanceof Boolean) {
                        cell.setCellValue((Boolean) object);
                    } else {
                        if (object != null) {
                            cell.setCellValue(object.toString());
                        }
                    }
                }
                index++;
            }
            try {
                outputStream = new FileOutputStream(filepath);
                workbook.write(outputStream);
                success = true;
            } catch (IOException e) {
                // 创建失败
            } finally {
                try {
                    if (outputStream != null) {
                        outputStream.close();
                    }
                    workbook.close();
                } catch (IOException ignored) {
                    // 流关闭失败
                }
            }
            return success;
        }
    }

    /**
     * 设置格式
     */
    private static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();
        // 标题样式
        XSSFCellStyle titleStyle = (XSSFCellStyle) wb.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER); // 水平对齐
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直对齐
        titleStyle.setLocked(true); // 样式锁定
        titleStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);
        titleFont.setFontName("微软雅黑");
        titleStyle.setFont(titleFont);
        styles.put("title", titleStyle);
        // 文件头样式
        XSSFCellStyle headerStyle = (XSSFCellStyle) wb.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex()); // 前景色
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 颜色填充方式
        headerStyle.setWrapText(true);
        headerStyle.setBorderRight(BorderStyle.THIN); // 设置边界
        headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        Font headerFont = wb.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        titleFont.setFontName("微软雅黑");
        headerStyle.setFont(headerFont);
        styles.put("header", headerStyle);
        // 字体样式
        Font cellStyleFont = wb.createFont();
        cellStyleFont.setFontHeightInPoints((short) 12);
        cellStyleFont.setColor(IndexedColors.BLUE_GREY.getIndex());
        cellStyleFont.setFontName("微软雅黑");
        // 正文样式A
        XSSFCellStyle cellStyleA = (XSSFCellStyle) wb.createCellStyle();
        cellStyleA.setAlignment(HorizontalAlignment.CENTER); // 居中设置
        cellStyleA.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleA.setWrapText(true);
        cellStyleA.setBorderRight(BorderStyle.THIN);
        cellStyleA.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setBorderLeft(BorderStyle.THIN);
        cellStyleA.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setBorderTop(BorderStyle.THIN);
        cellStyleA.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setBorderBottom(BorderStyle.THIN);
        cellStyleA.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleA.setFont(cellStyleFont);
        styles.put("cellA", cellStyleA);
        // 正文样式B:添加前景色为浅黄色
        XSSFCellStyle cellStyleB = (XSSFCellStyle) wb.createCellStyle();
        cellStyleB.setAlignment(HorizontalAlignment.CENTER);
        cellStyleB.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleB.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        cellStyleB.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyleB.setWrapText(true);
        cellStyleB.setBorderRight(BorderStyle.THIN);
        cellStyleB.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setBorderLeft(BorderStyle.THIN);
        cellStyleB.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setBorderTop(BorderStyle.THIN);
        cellStyleB.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setBorderBottom(BorderStyle.THIN);
        cellStyleB.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleB.setFont(cellStyleFont);
        styles.put("cellB", cellStyleB);
        return styles;
    }
}