package util;

import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;

/**
 * エクセルに関する処理をまとめたユーティリティ
 *
 * @author user
 */
public class ExcelUtil {
    /** 日付型のフォーマット */
    public static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

    /**
     * 行、列のインデックスからエクセル表記の位置文字列を返します。
     *
     * @param row 行インデックス
     * @param column 列インデックス
     * @return 「A1」のような位置文字列
     */
    public static String toPositionStr(int row, int column) {
        String colStr = CellReference.convertNumToColString(column);
        return colStr + String.valueOf(row + 1);
    }

    /**
     * データの入力規則(セルのプルダウン)を設定します。
     *
     * @param destCell 入力規則を設定するセル
     * @param range プルダウンに表示する値を入力した、「Sheet1!A1:B2」のような文字列
     */
    public static void createListConstraint(Cell destCell, String range) {
        Sheet sheet = destCell.getSheet();
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint listConstraint = helper.createFormulaListConstraint(range);
        CellRangeAddressList addressList = new CellRangeAddressList(//
                destCell.getRowIndex(), destCell.getRowIndex(), destCell.getColumnIndex(),
                destCell.getColumnIndex());
        DataValidation dataValidation = helper.createValidation(listConstraint, addressList);
        sheet.addValidationData(dataValidation);
    }

    /**
     * セルの値を文字列で返します。
     *
     * @param cell 読み込むセル
     * @return セルの入力値
     */
    public static String readCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                // return "";
                return readMergedCellValue(cell);
            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_ERROR:
                return FormulaError.forInt(cell.getErrorCellValue()).getString();
            case Cell.CELL_TYPE_FORMULA:
                // return cell.getCellFormula();
                return getFormulaResult(cell);
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                }
                return String.valueOf(cell.getNumericCellValue());
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                // TODO LOGGER
                throw new RuntimeException("unexpected cell type.");
        }
    }

    /**
     * 数式の計算結果を返します。
     *
     * @param cell 読み込むセル
     * @return 数式の計算結果
     */
    public static String getFormulaResult(Cell cell) {
        CreationHelper helper = cell.getSheet().getWorkbook().getCreationHelper();
        FormulaEvaluator evaluator = helper.createFormulaEvaluator();
        CellValue result = evaluator.evaluate(cell);
        switch (result.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(result.getBooleanValue());
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                }
                return String.valueOf(cell.getNumericCellValue());
            case Cell.CELL_TYPE_STRING:
                return result.getStringValue();
            default:
                // TODO LOGGER
                throw new RuntimeException("unexpected cell type.");
        }
    }

    /**
     * 結合セルの値を返します。
     *
     * @param cell 結合セルに含まれるセル
     * @return 結合セルの入力値
     */
    public static String readMergedCellValue(Cell cell) {
        Sheet sheet = cell.getSheet();
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        for (CellRangeAddress address : sheet.getMergedRegions()) {
            if (address.isInRange(rowIndex, columnIndex)) {
                Row row = sheet.getRow(address.getFirstRow());
                Cell firstCell = row.getCell(address.getFirstColumn());
                switch (firstCell.getCellType()) {
                    case Cell.CELL_TYPE_BLANK:
                        return "";
                    case Cell.CELL_TYPE_BOOLEAN:
                        return String.valueOf(cell.getBooleanCellValue());
                    case Cell.CELL_TYPE_ERROR:
                        return FormulaError.forInt(cell.getErrorCellValue()).getString();
                    case Cell.CELL_TYPE_FORMULA:
                        return getFormulaResult(cell);
                    case Cell.CELL_TYPE_NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            return sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                        }
                        return String.valueOf(cell.getNumericCellValue());
                    case Cell.CELL_TYPE_STRING:
                        return cell.getStringCellValue();
                    default:
                        // TODO LOGGER
                        throw new RuntimeException("unexpected cell type.");
                }
            }
        }
        return "";
    }
}
