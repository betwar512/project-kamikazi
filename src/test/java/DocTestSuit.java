import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;



import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;

/**
 * @author A.H.Safaie
 */
public class DocTestSuit {

    public static final String FILE_PATH = "docfiles/";

    XWPFDocument document;


    @Test public void createDocWithTable() throws IOException {

        //Blank Document
         document = new XWPFDocument();
        FileOutputStream out = new FileOutputStream(new File(FILE_PATH + "create_table.docx"));
        // create paragraph
        createNewPragraph("This is title ","someone somewhere bla bla bla bla ");

        //create table
        XWPFTable table = document.createTable();


        addHeadersToTable(table,Arrays.asList("Header One","Header Two","Header three"));
        List<List<Object>> listArrays = Arrays.asList(
                Arrays.asList("col one, row one","col two, row one", BigDecimal.valueOf(10)),
                Arrays.asList("col one, row two","col two, row two","col three, row two"),
                Arrays.asList("col one, row three","col two, row three","col three, row three")
        );

        for (int i = 0; i < listArrays.size(); i++) {

            addRow(table.createRow(),listArrays.get(i));
        }


        document.write(out);
        out.close();
        System.out.println("create_table.docx written successully");
    }

    void createNewPragraph(String title , String text){
        //TODO
        XWPFParagraph paragraph = document.createParagraph();
        //Set Bold an Italic
        XWPFRun titleRun = paragraph.createRun();
        titleRun.setFontSize(20);
        titleRun.setBold(true);
        titleRun.setText(title);
        titleRun.addBreak();

        //Set text Position
        XWPFRun textRun = paragraph.createRun();
        textRun.setFontSize(16);
        textRun.setText(text);

    }

    /**
     *
     * @param table
     * @param listStrs
     */
    void addHeadersToTable(XWPFTable table ,List<Object> listStrs){

        XWPFTableRow headerRow = table.getRow(0);
        for (int i = 0; i < listStrs.size(); i++) {

            Object o = listStrs.get(i);

            String strValue ;

            if( o instanceof String ){

                strValue = (String) o;

            } else {

                strValue = o.toString();

            }
            XWPFTableCell cel;

            if(i==0){
                 cel = headerRow.getCell(0);
            } else
                 cel = headerRow.addNewTableCell();

           // cel.setColor("");

            cel.setText(strValue);
        }


    }

    /**
     *  Create cell
     * @param tableRow
     * @param listStrs
     */
    private void addRow(XWPFTableRow tableRow, List<Object> listStrs){

        for (int i = 0; i < listStrs.size(); i++) {

            Object o = listStrs.get(i);

            String strValue ;

            if( o instanceof String ){

                 strValue = (String) o;

            } else {

                strValue = o.toString();

            }

            tableRow.getCell(i).setText(strValue);
        }
    }

}
