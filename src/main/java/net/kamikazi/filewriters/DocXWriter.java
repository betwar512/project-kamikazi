package net.kamikazi.filewriters;

import com.sun.istack.internal.Nullable;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

/**
 * @author A.H.Safaie
 */
public class DocXWriter {

    public static final String FILE_PATH = "/tmp/";
    public static final String FILE_TYPE = ".decx";



    private XWPFDocument document;
    //  private  FileOutputStream out;

    private String fileName ;

    /**
     *
     * @param fileName pass fileName or random gets generated
     */
    public DocXWriter(String fileName) {

        if(fileName != null && !fileName.isEmpty()){
            this.fileName = fileName;
        } else {

            this.fileName = UUID.randomUUID().toString();

        }

        this.document = new XWPFDocument();

    }


    /**
     * Finalise document
     * @return generate document object XWPFDocument
     * @throws IOException
     */
    public XWPFDocument finaliseDocument() throws IOException {
        FileOutputStream out = new FileOutputStream(new File(FILE_PATH + fileName + FILE_TYPE));
        document.write(out);
        out.close();
        return document;

    }

    /**
     *
     * @return file Name without any postfix
     */
    public String getFileName(){
        return this.fileName;
    }

    /**
     *
     * @return getAbsloute path of your file
     */
    public String getAbslouteFilePath(){
        return FILE_PATH + this.fileName + FILE_TYPE;
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
     *add document header for this docx
     * @param documentTitle center title text
     * @param leftDate left side text date
     * @param rightText right side text header
     * @param color color of text
     */
    public void addDocHeader(String documentTitle , Date leftDate , String rightText , String color ) {

        XWPFHeaderFooterPolicy policy = document.getHeaderFooterPolicy();

        if(StringUtils.isEmpty(color)){
            color = "6aa3c1";
        }
        String tabBetween = "\t";


        StringBuilder strb = new StringBuilder();
        if(leftDate != null ) {
            strb.append(leftDate.toString())
                    .append(tabBetween);
        }

        strb.append(documentTitle).append(tabBetween);
        strb.append(rightText);


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

            XWPFParagraph    paragraph = listPr.isEmpty() ?
                    headerD.createParagraph() : headerD.getParagraphs().get(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            XWPFRun runTitle = paragraph.createRun();
            runTitle.setFontSize(25);
            runTitle.setBold(true);
            runTitle.setColor(color);
            runTitle.setText(strb.toString());
            runTitle.addBreak();
            runTitle.addBreak();
        }
    }

    /**
     *  Add pragraph with title to Document
     * @param title Title text , defualt size 18
     * @param text main text text size is title size - 4  defualt 14
     * @param titleFontSize  Size of title --> size of text will be size of title - 4  , Defualt value 18
     * @param titleColor hex value color for title , use toHexString if you want to use java awt color
     * @param textColor hex value color for text ,  use toHexString if you want to use java awt color
     */
    public void addPragraphWithTitle(@Nullable String title ,@Nullable String text, int titleFontSize ,@Nullable String titleColor ,@Nullable String textColor ){

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
            if(titleColor != null) {
                titleRun.setColor(titleColor);
            }
        }

        if(!StringUtils.isEmpty(text)) {
            //Set text Position
            XWPFRun textRun = paragraph.createRun();
            textRun.setFontSize(titleFontSize - 4);
            textRun.setText(text);
            if(textColor != null) {
                textRun.setColor(textColor);
            }
            textRun.addBreak();
            textRun.addBreak();
        }
    }

    /**
     *
     * Add Line breaker
     * @param numberOfBreaks add more than one line breaker
     */
    public void addLineBreak(int numberOfBreaks){
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun textRun = paragraph.createRun();
        if(numberOfBreaks > 0){
            for (int i = 0; i < numberOfBreaks; i++) {
                textRun.addBreak();
            }
        } else {
            textRun.addBreak();
        }

    }


    /**
     * Create pragraph with multi breaks run
     * @param list list of strings
     * @param boldIndex  index of run with bold
     */
    public void addMutiSetPragraph(List<String> list ,@Nullable List<Integer> boldIndex){
        if(list != null && !list.isEmpty()) {
            if (boldIndex == null) {
                boldIndex = new ArrayList<>();
            }
            XWPFParagraph paragraph = document.createParagraph();
            for (int i = 0; i < list.size(); i++) {
                setRun(paragraph.createRun(), "", 12, "", list.get(i), boldIndex.contains(i), true);
            }
        }
    }





    /**
     * Use run to create multi line pragraph and specify font details for each part
     * @param run run created from pragraph
     * @param fontFamily type of font can be empu
     * @param fontSize ignore if 0
     * @param colorRGB color string heax value without #
     * @param text value of run
     * @param bold use bold font
     * @param addBreak add line breaked
     */
    public void setRun(XWPFRun run , @Nullable String fontFamily , int fontSize ,@Nullable String colorRGB , String text , boolean bold , boolean addBreak) {

        if(fontFamily != null) {
            run.setFontFamily(fontFamily);
        }

        if(fontSize > 0) {
            run.setFontSize(fontSize);
        }
        if(colorRGB != null) {
            run.setColor(colorRGB);
        }

        run.setText(text);

        run.setBold(bold);

        if (addBreak) run.addBreak();
    }


    /*--------------------------------------------------------------------------------------------------------------------*/
    /*   								Tables											   					      */
    /*------------------------------------------------------------------------------------------------------------------*/


    /**
     * Create new table
     * @param setBorders send true to set all the borders
     * @return XWPFTable table
     */
    public  XWPFTable addNEwTable(boolean setBorders){
        XWPFTable table = document.createTable();
        //no border
        if(!setBorders) {
            table.getCTTbl().getTblPr().unsetTblBorders();
        }
        CTTblWidth width = table.getCTTbl().addNewTblPr().addNewTblW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(9072));


        return table;

    }


    /**
     *
     * @param table target table
     * @param listStrs list cell values
     * @param height height of header row , defualt 600
     * @param topBorder if set top border
     * @param bottomBorder to set bottom border
     */
    public void addHeadersToTable(XWPFTable table ,List<Object> listStrs , int height, boolean topBorder , boolean bottomBorder){

        if(height == 0 ){
            height = 600;
        }

        XWPFTableRow headerRow = table.getRow(0);
        headerRow.setHeight(height);
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

            if(bottomBorder) {
                borderCe.addNewBottom().setVal(STBorder.SINGLE);
            }

            if(topBorder) {
                borderCe.addNewTop().setVal(STBorder.SINGLE);
            }
            // background

            tcpr.addNewShd().setFill(toHexString(Color.LIGHT_GRAY));

            XWPFParagraph pr = cel.getParagraphs().isEmpty()? cel.addParagraph() :  cel.getParagraphs().get(0);
            pr.setAlignment(ParagraphAlignment.CENTER);
            setRun(pr.createRun(),"",0,"",strValue.toUpperCase(),true,false);
            // cel.setText(strValue.toUpperCase());
        }
    }

    /**
     * Add new row to table
     * @param table add row to this table
     * @param listStrs list of cells values
     * @param rowHeight height of row
     */
    public void addRow(XWPFTable table , List<Object> listStrs , int rowHeight){
        XWPFTableRow tableRow = table.createRow();
        if(rowHeight == 0 ){
            rowHeight = 1440;
        }
        tableRow.setHeight(rowHeight);
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


    public XWPFTableRow addEmptyRow(XWPFTable table ,int rowHeight){
        XWPFTableRow tableRow = table.createRow();
        if(rowHeight == 0 ){
            rowHeight = 1440;
        }

        if(rowHeight > 0) {
            tableRow.setHeight(rowHeight);
        }
        return tableRow;
    }

    /**
     *
     * @param tableRow
     * @param index
     * @return
     */
    public XWPFTableCell addCellToRow( XWPFTableRow tableRow ,int index ,boolean addBorderBottm,String strValue ){

        XWPFTableCell cel = tableRow.getCell(index);
        XWPFParagraph pr = cel.getParagraphs().isEmpty()? cel.addParagraph() :  cel.getParagraphs().get(0);
        pr.setAlignment(ParagraphAlignment.CENTER);

        CTTc    ctt = cel.getCTTc();
        CTTcPr tcpr = ctt.addNewTcPr();

        // border
        CTTcBorders borderCe = tcpr.addNewTcBorders();
        if(addBorderBottm) {
            borderCe.addNewBottom().setVal(STBorder.SINGLE);
        }
        if(!StringUtils.isEmpty(strValue)){
            cel.setText(strValue);
        }

        return cel;
    }




    /**
     *
     * @param colour
     * @return
     * @throws NullPointerException
     */
    public static String toHexString(Color colour) throws NullPointerException {
        String hexColour = Integer.toHexString(colour.getRGB() & 0xffffff);
        if (hexColour.length() < 6) {
            hexColour = "000000".substring(0, 6 - hexColour.length()) + hexColour;
        }
        return  hexColour;
    }

}


