package util;

import static org.hamcrest.CoreMatchers.*;
import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.junit.experimental.runners.Enclosed;
import org.junit.runner.RunWith;

/**
 * {@link ExcelUtil}のテストクラス
 *
 * @author user
 */
@RunWith(Enclosed.class)
@SuppressWarnings("javadoc")
public class ExcelUtilTest {
    public static class エクセルファイル不要 {
        @Test
        public void toPositionStrはセル番地の文字列を返す() throws Exception {
            String actual = ExcelUtil.toPositionStr(26, 26);
            String expected = "AA27";
            assertThat(actual, is(expected));
        }
    }

    public static class エクセルファイル入出力 {
        private static final String INPUT_NAME = "test.xls";
        private static final String SRC_SHEET_NAME = "入力シート";
        private static final String DEST_SHEET_NAME = "出力シート";
        Workbook wb;

        @Test
        public void createListConstraintは指定セルにプルダウンを設定する() throws Exception {
            // TODO testケース
        }

        @Before
        public void setUp() throws Exception {
            wb = WorkbookFactory.create(ExcelUtilTest.class.getClassLoader().getResourceAsStream(INPUT_NAME));
        }

        @After
        public void tearDown() throws Exception {
            wb.close();
        }
    }
}
