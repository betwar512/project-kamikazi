package net.kamikazi.filewriters;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;
import java.util.UUID;

public class DocXWriter {


    private static final String FILE_PATH = "temp/";


    private XWPFDocument document;
    private FileOutputStream out;

    private String fileName ;


    public DocXWriter(String fileName) throws FileNotFoundException {

        if(fileName != null && !fileName.isEmpty()){
            this.fileName = fileName;
        } else {

            this.fileName = UUID.randomUUID().toString();

        }

        this.document = new XWPFDocument();
        this.out = new FileOutputStream(new File(FILE_PATH + this.fileName + ".docx"));
    }


    /*--------------------------------------------------------------------------------------------------------------------*/
    /*   						Document content 											   					      */
    /*------------------------------------------------------------------------------------------------------------------*/


    /**
     * Setup page size for thie document
     */
    public void setPageSize(int pageWidth , int pageHeight){
        CTBody body = document.getDocument().getBody();
        CTSectPr section = body.getSectPr();
        if(section == null){
            section = body.addNewSectPr();
            section.addNewPgSz();
        }
        CTPageSz pageSize = section.getPgSz();
        pageSize.setW(BigInteger.valueOf(pageWidth));
        pageSize.setH(BigInteger.valueOf(pageHeight));

    }



    /**
     *Add header to the document
     * @param documentTitle
     */
    public void addDocHeader(String documentTitle) {

        XWPFHeaderFooterPolicy policy = document.getHeaderFooterPolicy();

        //in an empty document always will be null
        if(policy == null){
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
     *  Add pragraph with title to Document
     * @param title Title text , text size is title size - 4  defualt 14
     * @param text main text
     * @param titleFontSize  Size of title --> size of text will be size of title - 4  , Defualt value 18
     */
    public void addPragraphWithTitle(String title , String text , int titleFontSize){

        if(titleFontSize == 0 ){
            titleFontSize = 18;
        }


        XWPFParagraph paragraph = document.createParagraph();
        if(!StringUtils.isEmpty(title)) {
            //Set Bold an Italic
            XWPFRun titleRun = paragraph.createRun();
            titleRun.setFontSize(titleFontSize);
            titleRun.setBold(true);
            titleRun.setText(title);
            titleRun.addBreak();
        }

        //Set text Position
        XWPFRun textRun = paragraph.createRun();
        textRun.setFontSize(titleFontSize - 4 );
        textRun.setText(text);
        textRun.addBreak();
        textRun.addBreak();

    }


    /**
     * Create pragraph with multi breaks run
     * @param list list of strings
     * @param boldIndex  index of run with bold
     */
    public void addMutiSetPragraph(List<String> list ,List<Integer> boldIndex){

        XWPFParagraph paragraph = document.createParagraph();
        for (int i = 0; i < list.size(); i++) {
            setRun(paragraph.createRun(),"",12,"",list.get(i),boldIndex.contains(i),true);
        }
    }


    /**
     *
     * @param run
     * @param fontFamily
     * @param fontSize
     * @param colorRGB
     * @param text
     * @param bold
     * @param addBreak
     */
    public void setRun(XWPFRun run , String fontFamily , int fontSize , String colorRGB , String text , boolean bold , boolean addBreak) {

        run.setFontFamily(fontFamily);
        if(fontSize > 0)
            run.setFontSize(fontSize);

        run.setColor(colorRGB);

        run.setText(text);

        run.setBold(bold);

        if (addBreak) run.addBreak();
    }


    /*--------------------------------------------------------------------------------------------------------------------*/
    /*   								Tables											   					      */
    /*------------------------------------------------------------------------------------------------------------------*/

    /**
     *
     * @param table
     * @param listStrs
     */
    public void addHeadersToTable(XWPFTable table ,List<Object> listStrs){

        XWPFTableRow headerRow = table.getRow(0);
        headerRow.setHeight(600);
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
            CTTc ctt = cel.getCTTc();
            CTTcPr tcpr = ctt.addNewTcPr();

            // border
            CTTcBorders borderCe = tcpr.addNewTcBorders();

            borderCe.addNewBottom().setVal(STBorder.SINGLE);

            borderCe.addNewTop().setVal(STBorder.SINGLE);

            // background

            tcpr.addNewShd().setFill(toHexString(Color.LIGHT_GRAY));

            XWPFParagraph pr = cel.getParagraphs().isEmpty()? cel.addParagraph() :  cel.getParagraphs().get(0);
            pr.setAlignment(ParagraphAlignment.CENTER);
            setRun(pr.createRun(),"",0,"",strValue.toUpperCase(),true,false);
            cel.setText(strValue.toUpperCase());
        }
    }

    /**
     *  Create cell
     * @param tableRow
     * @param listStrs
     */
    public void addRow(XWPFTableRow tableRow, List<Object> listStrs){
        tableRow.setHeight(1440);
        for (int i = 0; i < listStrs.size(); i++) {

            Object o = listStrs.get(i);

            String strValue ;

            if( o instanceof String ){

                strValue = (String) o;

            } else {

                strValue = o.toString();

            }
            XWPFTableCell cel = tableRow.getCell(i);
            XWPFParagraph pr = cel.getParagraphs().isEmpty()? cel.addParagraph() :  cel.getParagraphs().get(0);
            pr.setAlignment(ParagraphAlignment.CENTER);
            cel.setText(strValue);

            CTTc    ctt = cel.getCTTc();
            CTTcPr tcpr = ctt.addNewTcPr();

            // border
            CTTcBorders borderCe = tcpr.addNewTcBorders();

            borderCe.addNewBottom().setVal(STBorder.SINGLE);
        }
    }



    /**
     *
     * @param colour
     * @return
     * @throws NullPointerException
     */
    String toHexString(Color colour) throws NullPointerException {
        String hexColour = Integer.toHexString(colour.getRGB() & 0xffffff);
        if (hexColour.length() < 6) {
            hexColour = "000000".substring(0, 6 - hexColour.length()) + hexColour;
        }
        return  hexColour;
    }

}
