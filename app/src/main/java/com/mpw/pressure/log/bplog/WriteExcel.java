package com.mpw.pressure.log.bplog;

/**
 * Created by malith on 8/5/17.
 */

import android.content.Context;
import android.os.Environment;
import android.widget.Toast;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class WriteExcel {

    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;
    private Context context;


    public void write(String _inputFile, Context con) throws IOException, WriteException {
        this.context = con;
        this.inputFile = _inputFile;


        File file = new File(context.getExternalFilesDir(null), inputFile);

        if (canWriteOnExternalStorage()) {

            WorkbookSettings wbSettings = new WorkbookSettings();

            wbSettings.setLocale(new Locale("en", "EN"));

            WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
            workbook.createSheet("Report", 0);
            WritableSheet excelSheet = workbook.getSheet(0);
            createLabel(excelSheet);
            createContent(excelSheet);

            workbook.write();
            workbook.close();
        } else {
            Toast.makeText(context, "Unable to Write to External Storage.", Toast.LENGTH_SHORT).show();
        }
    }

    public static boolean canWriteOnExternalStorage() {
        // get the state of your external storage
        String state = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(state)) {
            return true;
        }
        return false;
    }

    private void createLabel(WritableSheet sheet)
            throws WriteException {
        // Lets create a times font
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        // Define the cell format
        times = new WritableCellFormat(times10pt);
        // Lets automatically wrap the cells
        times.setWrap(true);

        // create create a bold font with unterlines
        WritableFont times10ptBoldUnderline = new WritableFont(
                WritableFont.TIMES, 10, WritableFont.BOLD, false,
                UnderlineStyle.SINGLE);
        timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
        // Lets automatically wrap the cells
        timesBoldUnderline.setWrap(true);

        CellView cv = new CellView();
        cv.setFormat(times);
        cv.setFormat(timesBoldUnderline);
        cv.setAutosize(true);

        // Write a few headers
        addCaption(sheet, 0, 0, "SYS(mmhg)");
        addCaption(sheet, 1, 0, "DIA(mmhg)");
        addCaption(sheet, 2, 0, "PLUS/min");
        addCaption(sheet, 3, 0, "Date");
        addCaption(sheet, 4, 0, "Time");


    }

    private void createContent(WritableSheet sheet) throws WriteException {
        // Write a few number
        for (int i = 1; i < 10; i++) {
            // First column
            addNumber(sheet, 0, i, i * 10);
            // Second column
            addNumber(sheet, 1, i, i * 5);
            // Third column
            addNumber(sheet, 2, i, 72);
            // Fourth column
            addString(sheet, 3, i, "2017/08/06");
            // Fifth column
            addString(sheet, 4, i, "11:10:24");
        }
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s)
            throws WriteException {
        Label label;
        label = new Label(column, row, s, timesBoldUnderline);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row,
                           Integer integer) throws WriteException {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addString(WritableSheet sheet, int column, int row, String s)
            throws WriteException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }
}