package com.helen.vendor;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public abstract class XLSXReaderHelper extends DefaultHandler {

    //存储单元格table
    private SharedStringsTable sharedStringsTable;

    //读取到内容存放处
    private String content;

    //判断是否字符串，用于转字符串
    private boolean isString;

    //sheet索引
    private int sheetIndex = -1;

    //单行数据集
    private List<String> rowList = new ArrayList<String>();

    //当前行
    private int curRow = 0;

    //当前列
    private int curCol = 0;

    //上次操作的列（用于比对头部是否有空单元格）
    private int preCol = 0;


    //接口暴露方法
    public abstract void fetchRows(int sheetIndex, int curRow, List<String> rowList);

    /**
     * 执行方法
     *
     * @param filename
     * @param sheetId
     * @throws OpenXML4JException
     * @throws IOException
     * @throws SAXException
     */
    public void fetchOneSheet(String filename, int sheetId) throws OpenXML4JException, IOException, SAXException {
        sheetIndex = sheetId;
        OPCPackage opcPackage = OPCPackage.open(filename);
        XSSFReader xssfReader = new XSSFReader(opcPackage);
        SharedStringsTable stringsTable = xssfReader.getSharedStringsTable();
        XMLReader xmlReader = this.fetchSheetParser(stringsTable);
        InputStream inputStream = xssfReader.getSheet("rId" + sheetId);
        InputSource inputSource = new InputSource(inputStream);
        xmlReader.parse(inputSource);
        inputStream.close();
        opcPackage.close();
    }

    /**
     * 开始读取元素
     *
     * @param uri
     * @param localName
     * @param qName
     * @param attributes
     * @throws SAXException
     */
    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        //判断格式是否属于单元格
        if (qName.equals("c")) {
            //单元格格式
            String cellType = attributes.getValue("t");
            //单元格列索引字符串
            String colIndexStr = attributes.getValue("r");
            curCol = getColIndex(colIndexStr);
            isString = false;
            if (cellType != null && cellType.equals("s")) {
                isString = true;
            }
        }
        content = "";
    }

    /**
     * 结束读取元素
     *
     * @param uri
     * @param localName
     * @param qName
     * @throws SAXException
     */
    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        //字符串转型
        if (isString) {
            try {
                int index = Integer.parseInt(content);
                content = new XSSFRichTextString(sharedStringsTable.getEntryAt(index)).toString();
                isString=false;
            } catch (Exception e) {

            }
        }

        //v为单元格的值
        if (qName.equals("v")) {
            String cellValue = content.trim();
            cellValue = cellValue.equals("") ? null : cellValue;
            int cols = curCol - preCol;
            if (cols > 1) {
                for (int i = 0; i < cols - 1; i++) {
                    rowList.add(preCol, null);
                }
            }
            preCol = curCol;
            rowList.add(curCol - 1, cellValue);
        } else if (qName.equals("row")) {
            try {
                fetchRows(sheetIndex, curRow, rowList);
            } catch (Exception e)
            {
                e.printStackTrace();
            }

            rowList.clear();
            curRow++;
            curCol = 0;
            preCol = 0;
        }
    }

    /**
     * 获取单元格内容
     *
     * @param ch
     * @param start
     * @param length
     * @throws SAXException
     */
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        content += new String(ch, start, length);
    }

    /**
     * 获取XMLReader对象
     *
     * @param sharedStringsTable
     * @return
     * @throws SAXException
     */
    private XMLReader fetchSheetParser(SharedStringsTable sharedStringsTable) throws SAXException {
        XMLReader xmlReader = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sharedStringsTable = sharedStringsTable;
        xmlReader.setContentHandler(this);
        return xmlReader;
    }

    /**
     * 获取列索引位置
     * 使用字母的ASCII码进行转换
     *
     * @param colStr
     * @return
     */
    private int getColIndex(String colStr) {
        colStr = colStr.replaceAll("[^A-Z]", "");
        byte[] colByte = colStr.getBytes();
        int colByteLength = colByte.length;
        float num = 0;
        for (int i = 0; i < colByteLength; i++) {
            num += (colByte[i] - 'A' + 1) * Math.pow(26, colByteLength - i - 1);
        }
        return (int) num;
    }
}
