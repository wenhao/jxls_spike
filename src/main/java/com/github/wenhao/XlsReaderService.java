package com.github.wenhao;

import com.github.wenhao.domain.CheckItem;
import com.google.common.io.Resources;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jxls.reader.ReaderBuilder;
import org.xml.sax.SAXException;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class XlsReaderService {

    public static void main(String[] args) throws Exception {
        URL resource = Resources.getResource("demo.xls");
        List<CheckItem> checkItems = read(resource, "analysis.xml");
        System.out.println(checkItems.size());
    }

    private static List<CheckItem> read(URL resource, String templateFile) throws IOException, InvalidFormatException, SAXException {
        Workbook workbook = WorkbookFactory.create(resource.openStream());
        workbook.sheetIterator().forEachRemaining(XlsReaderService::unMergeCells);
        workbook.write(new FileOutputStream(resource.getFile()));
        List<CheckItem> checkItems = new ArrayList<>();
        final Map<Object, Object> map = new HashMap<>();
        map.put("checkItems", checkItems);
        ReaderBuilder.buildFromXML(Resources.getResource(templateFile).openStream())
                .read(resource.openStream(), map);
        return checkItems;
    }

    public static void unMergeCells(Sheet sheet) {
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            String value = sheet.getRow(region.getFirstRow())
                    .getCell(region.getFirstColumn())
                    .getStringCellValue();
            sheet.removeMergedRegion(i);
            sheet.rowIterator().forEachRemaining(row -> row.cellIterator().forEachRemaining(cell -> {
                if (region.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                    cell.setCellValue(value);
                }
            }));
        }
    }
}
