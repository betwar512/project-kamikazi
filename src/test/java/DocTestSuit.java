import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

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

        addDocTitle("My generated Doc");

        // create paragraph
        createNewParagraph("This is title ","someone somewhere bla bla bla bla ");

        createNewParagraph("","new text this is second line with some new content added no title ");

        createNewParagraph("Another one","new text this is second line with some new content added no title ");

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
        System.out.println("create_table.docx written successfully");
    }

    /**
     *
     * @param title
     * @param text
     */
    void createNewParagraph(String title , String text){

        XWPFParagraph paragraph = document.createParagraph();
        if(!StringUtils.isEmpty(title)) {
            //Set Bold an Italic
            XWPFRun titleRun = paragraph.createRun();
            titleRun.setFontSize(18);
            titleRun.setBold(true);
            titleRun.setText(title);
            titleRun.addBreak();
        }

        //Set text Position
        XWPFRun textRun = paragraph.createRun();
        textRun.setFontSize(14);
        textRun.setText(text);
        textRun.addBreak();
        textRun.addBreak();

    }



    /**
     *
     * @param documentTitle
     */
    void addDocTitle(String documentTitle){

        XWPFHeaderFooterPolicy policy = document.getHeaderFooterPolicy();
        //in an empty document always will be null
        if(policy==null){
            CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
            policy = new  XWPFHeaderFooterPolicy( document, sectPr );
        }
        if (policy.getDefaultHeader() == null && policy.getFirstPageHeader() == null

                && policy.getDefaultFooter() == null) {
            // Need to create some new headers
            // The easy way, gives a single empty paragraph
            XWPFHeader headerD = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

            List<XWPFParagraph> listPr = headerD.getParagraphs();


            if(listPr.isEmpty()){

                XWPFParagraph paragraph = headerD.createParagraph();

                paragraph.setAlignment(ParagraphAlignment.CENTER);

                XWPFRun runTitle = paragraph.createRun();

                runTitle.setFontSize(25);

                runTitle.setBold(true);

                runTitle.setColor("6aa3c1");

                runTitle.setText(documentTitle);

                runTitle.addBreak();

                runTitle.addBreak();

            } else {
                headerD.getParagraphs().get(0).createRun().setText(documentTitle);
            }
        }
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
