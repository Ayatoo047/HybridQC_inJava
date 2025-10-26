package com.labysolutions.j_hybridqc;


import org.apache.commons.math3.exception.ZeroException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;


import java.awt.Color;
import java.io.*;
import java.time.Duration;
import java.time.Instant;
import java.util.*;

public class ExcelLogicClass {

    private Sheet sheet;
    private String curr_parent;
    private String next_parent;
    private String curr_value;
    private String next_value;

    private final int min_missing_percentage = 20;
    private final int min_perc_polymorphic = 20;
    private final int min_perc_hybridity = 50;
    private String saveas;

    int track_parent = 0;
    int parent_polymophic = 0;
    int parent_missing = 0;

    int parent1_het = 0;
    int parent2_het = 0;

    int parent1_missing = 0;
    int parent2_missing = 0;

    int marker_col_start = 14;
    int no_markers;
    Workbook workbook;


    private final int max_perc_het = 20;
    private final String title = "Results";
    private final HashMap<String, Object> dict_skip = new HashMap<>();
    private final ArrayList<Object[]> marker_success = new ArrayList<Object[]>();

    private int undefined_missing_data = 0;
    private int undefined_unpolymorphic_parent = 0;
    private int undefined_parent_het = 0;

    private double percent_polymorph = 0;

    private int TRUE = 0;
    private int FAILED = 0;


    final private int polymophic_col = 2;
    final private int perc_polymorphic_col = 3;

    final private int no_parentHet_col = 4;
    final private int perc_parentHet_col = 5;

    final private int no_outcrossing_col = 6;
    final private int perc_outcrossing_col = 7;

    final private int trueF1_col = 8;
    final private int missing_col = 9;
    final private int perc_missing_col = 10;
    final private int hybridity_col = 11;
    final private int status_col = 12;
    final private int parent_col = 13;
    private final int marker_start_col = 14;
    private int track_parent_on_F1 = 0;


    double percHet;


    Set<String> heterozyte_values = new HashSet<>(Arrays.asList("A:T", "A:C", "A:G", "T:A", "T:C", "T:G", "C:A",
            "C:T", "C:G", "G:A", "G:T", "G:C"));

    Set<String> homogousAndMissingSet = new HashSet<>(Arrays.asList(
            "G:G", "A:A", "C:C", "T:T", "Uncallable", "?"
    ));
    Set<String> homogous_value = new HashSet<>(Arrays.asList("G:G", "A:A", "C:C", "T:T"));
    Set<String> missing_values = new HashSet<>(Arrays.asList("Uncallable", "?"));

    Color green_color = new Color(0, 255, 0);
    Color blue_color = new Color(0, 123, 255);
    Color red_color = new Color(255, 0, 0);
    Color orange_color = new Color(255, 165, 0);
    Color magenta_color = new Color(255, 0, 255);
    Color grey_color = new Color(192, 192, 192);



    XSSFColor green = new XSSFColor(green_color, new DefaultIndexedColorMap());
    XSSFColor blue = new XSSFColor(blue_color, new DefaultIndexedColorMap());
    XSSFColor red = new XSSFColor(red_color, new DefaultIndexedColorMap());
    XSSFColor orange = new XSSFColor(orange_color, new DefaultIndexedColorMap());
    XSSFColor magenta = new XSSFColor(magenta_color, new DefaultIndexedColorMap());
    XSSFColor grey = new XSSFColor(grey_color, new DefaultIndexedColorMap());



    XSSFCellStyle redstyle;
    XSSFCellStyle greenstyle;
    XSSFCellStyle bluestyle;
    XSSFCellStyle orangestyle;
    XSSFCellStyle magentastyle;
    XSSFCellStyle greystyle;
    XSSFCellStyle style;

    ArrayList<XSSFCellStyle> stylesHash = new ArrayList<XSSFCellStyle>();

    private Cell curr_cell;
    private Cell next_cell;

    private Row curr_row;
    private Row next_row;

    private int no_outcross;
    private int missiing = 0;
    private int true_values = 0;
    int maxColumn = 0;


//    greencell1.setCellStyle(style);

    private void createStyledCell(XSSFCellStyle style, XSSFColor color) {
//        XSSFCellStyle newStyle = (XSSFCellStyle) workbook.createCellStyle(); // Create new style each time
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        return newStyle;
    }


    public void start() throws Exception {

        Instant start = Instant.now();



        String filename = "/Users/m1/Documents/J_HybridQC/src/main/resources/1000 F1 + parent + 22SNPS.xlsx";
        this.getSheet(filename);

        this.createStyledCell(redstyle, red);
        this.createStyledCell(greenstyle, green);
        this.createStyledCell(bluestyle, blue);
        this.createStyledCell(orangestyle, orange);
        this.createStyledCell(magentastyle, magenta);
        this.createStyledCell(greystyle, grey);

        this.checkPolymorphicParent();
//        this.checkPolymorphicParentBackup();
        this.createStatHeaders();
        this.f1Check();
        this.createPieChart();
        this.createBarChart();
        this.save();


        Instant end = Instant.now();
        Duration timeElapsed = Duration.between(start, end);

        System.out.println("Function took: " + timeElapsed.toMillis() + " ms");
    }

    private void createBarChart() {

        System.out.println("===================");
        var efficiencySheet = workbook.createSheet("Efficiency");
        Row header = efficiencySheet.createRow(0);
        header.createCell(0).setCellValue("SNPs");
        header.createCell(1).setCellValue("#Parent Combination");
        header.createCell(2).setCellValue("Polymorphism Frequency");
        header.createCell(3).setCellValue("Marker Efficiency (%)");

        ArrayList<Object[]> barData = new ArrayList<Object[]>();

        for (Object[] rowData : marker_success){
            int marker_index = (int) rowData[0];
            String colLetter = (String) rowData[1];
            int num_of_success = (int) rowData[2];

            int rowNumber = CellReference.convertColStringToIndex(colLetter);
            String SNPvalue = sheet.getRow(rowNumber).getCell(0).getStringCellValue();



            double efficiency_percentage = (double) num_of_success /track_parent * 100;
//            efficiencySheet.getRow();
            Object[] data = {SNPvalue, track_parent, num_of_success, efficiency_percentage};

            barData.add(data);
        }


        int rowIdx = 1;
        for (Object[] rowData : barData) {
            Row row = efficiencySheet.createRow(rowIdx++);
            row.createCell(0).setCellValue((String) rowData[0]);
            row.createCell(1).setCellValue((Integer) rowData[1]);
            row.createCell(2).setCellValue((Integer) rowData[2]);
            row.createCell(3).setCellValue((Double) rowData[3]);
        }
//        efficiencySheet.

        XSSFDrawing drawing = (XSSFDrawing) efficiencySheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 5, 1, 15, 20);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("SNPs Performance");
        chart.setTitleOverlay(false);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Markers");
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("Marker Efficiency (%)");

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) efficiencySheet,
                new CellRangeAddress(1, barData.size(), 0, 0)
        );
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) efficiencySheet,
                new CellRangeAddress(1, barData.size(), 3, 3)
        );

        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
        XDDFChartData.Series series = data.addSeries(categories, values);
        series.setTitle("Marker Efficiency (%)", null);

        chart.plot(data);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);
    }

    private void save() {

        try (FileOutputStream fileOut = new FileOutputStream("/Users/m1/Documents/HybridQC.Out_TestFile2.xlsx")) {
            workbook.write(fileOut);
            workbook.close();  // optional but good practice
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


//    private void insertColumns(int startCol, int numCols) {
//        int lastRowNum = this.sheet.getLastRowNum();
//
//        for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
//            Row row = sheet.getRow(rowNum);
//            if (row == null) continue;
//
//            // Shift cells to the right
//            for (int colNum = row.getLastCellNum() - 1; colNum >= startCol - 1; colNum--) {
//                Cell oldCell = row.getCell(colNum);
//                Cell newCell = row.createCell(colNum + numCols);
//
//                if (oldCell != null) {
//                    cloneCell(oldCell, newCell);
//                    row.removeCell(oldCell);
//                }
//            }
//        }
//    }

    private void insertColumns(int startCol, int numCols) {
        int lastRowNum = sheet.getLastRowNum();

        for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) continue;

            // Shift existing cells to the right
            short lastCellNum = row.getLastCellNum();
            if (lastCellNum < 0) continue;

            for (int colNum = lastCellNum - 1; colNum >= startCol - 1; colNum--) {
                Cell oldCell = row.getCell(colNum);
                Cell newCell = row.createCell(colNum + numCols);

                if (oldCell != null) {
                    cloneCell(oldCell, newCell);
                    row.removeCell(oldCell);
                }
            }

            // ✅ Create empty cells in the newly inserted range
            for (int colNum = startCol - 1; colNum < startCol - 1 + numCols; colNum++) {
                if (row.getCell(colNum) == null) {
                    row.createCell(colNum);
                }
            }
        }
    }


    private void cloneCell(Cell oldCell, Cell newCell) {
        newCell.setCellStyle(oldCell.getCellStyle());

        switch (oldCell.getCellType()) {
            case STRING -> newCell.setCellValue(oldCell.getStringCellValue());
            case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
            case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
            case FORMULA -> newCell.setCellFormula(oldCell.getCellFormula());
            default -> { }
        }
    }

    public void getSheet(String filename) throws IOException {
//        filename = "/Users/m1/Documents/J_HybridQC/src/main/resources/HybridQC.TestFile.xlsx";
//        filename = "/Users/m1/Documents/J_HybridQC/src/main/resources/HybridQC.TestFile.xlsx";

        try (var fis = new FileInputStream(new File(filename))) {
//            Workbook workbook;

            if (filename.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (filename.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("Invalid file type: " + filename);
            }

//            this.sheet = workbook.getSheet(String.valueOf(0));    todo: get sheet by name
            this.sheet = workbook.getSheetAt(0);
            if (this.sheet == null) {
                throw new IllegalStateException("❌ No sheet found in workbook!");
            }
            this.insertColumns(3, 11);

            for (Row row : sheet) {
                if (row.getLastCellNum() > maxColumn) {
                    maxColumn = row.getLastCellNum();
                }
            }
            this.no_markers = maxColumn - this.marker_col_start;

            redstyle = (XSSFCellStyle) workbook.createCellStyle();
            greenstyle = (XSSFCellStyle) workbook.createCellStyle();
            bluestyle = (XSSFCellStyle) workbook.createCellStyle();
            orangestyle = (XSSFCellStyle) workbook.createCellStyle();
            magentastyle = (XSSFCellStyle) workbook.createCellStyle();
            greystyle = (XSSFCellStyle) workbook.createCellStyle();

//            XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();

            style = (XSSFCellStyle) workbook.createCellStyle();

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    public void checkPolymorphicParent() {

        for (int i = 1; i < sheet.getLastRowNum(); i++) { // i=1 means start from row 2 (0-based index)
            Row row1 = sheet.getRow(i);
            Row row2 = sheet.getRow(i + 1); // next row

            if (row1 == null || row2 == null) continue;


            this.curr_row = row1;
            this.next_row = row2;


//            // Handle cell values like Python's value/value2
//            String val1 = getCellValue(cell1);
//            String val2 = getCellValue(cell2);






//            this.curr_value = val1;
//            this.next_value = val2;

            if (this.isParentPair()) {

                this.track_parent += 1;

                // todo: color the cells
//                    track_parent = 0;
                parent_polymophic = 0;
                parent_missing = 0;

                parent1_het = 0;
                parent2_het = 0;

                parent1_missing = 0;
                parent2_missing = 0;

//                    setStyleColor(red);
                curr_row.getCell(parent_col).setCellStyle(greenstyle);
                next_row.getCell(parent_col).setCellStyle(greenstyle);
                // working till this point

                for (int k = this.marker_col_start; k < row1.getLastCellNum(); k++) {
                    int marker_index = k - this.marker_col_start;

                    String colLetter = CellReference.convertNumToColString(k);


                    curr_value = curr_row.getCell(k).getStringCellValue();
                    next_value = next_row.getCell(k).getStringCellValue();

                    curr_cell = curr_row.getCell(k);
                    next_cell= next_row.getCell(k);

                    String val1 = getCellValue(curr_cell);
                    String val2 = getCellValue(next_cell);

                    String dict_value;

                    if (marker_success.size() < no_markers){
                        Object[] marker_success_value = {marker_index, colLetter, 0};
                        marker_success.add(marker_success_value);
                    }


                    if (!val1.equals(val2) && (this.isNotMissing())) {

                        this.colorAndSkipIfParentHet(colLetter);
//
                        // marker success
                        Object[] initial_marker_success = marker_success.get(marker_index);
                        int num_of_success = (((int) (initial_marker_success[2])) + 1);
                        Object[] new_marker_succ_value = {marker_index, colLetter, num_of_success };
                        marker_success.set(marker_index, new_marker_succ_value);
                    }
                    else if (this.bothMissing()) {

//                            setStyleColor(red);

//                            String colLetter = CellReference.convertNumToColString(k);
//                            var dict_value = "parent" + String.valueOf(track_parent) + String.valueOf(curr_row.getRowNum());
                        dict_value = "parent" + track_parent + colLetter;
                        dict_skip.put(dict_value, "Skip");
                        curr_cell.setCellStyle(redstyle);
                        next_cell.setCellStyle(redstyle);

                        parent1_missing += 1;
                        parent2_missing += 1;
                        parent_missing += 1;
                    }
                    else {
                        // skip
//

                        dict_value = "parent" + track_parent + colLetter;

                        dict_skip.put(dict_value, "Skip");
                        if (!homogousAndMissingSet.contains(curr_value) && !homogousAndMissingSet.contains(next_value)) {
//
                            setStyleColor(blue);

                            curr_cell.setCellStyle(bluestyle);
                            next_cell.setCellStyle(bluestyle);

                            this.parent1_het += 1;
                            this.parent2_het += 1;
                        } else {

                            setStyleColor(red);

                            curr_cell.setCellStyle(redstyle);
                            next_cell.setCellStyle(redstyle);
                        }

                    }
                }
//
                this.setPolymorphicHybridityValues();
                this.colorParentGrey();



            }
        }

    }

    public void checkPolymorphicParentBackup() {

        for (int i = 1; i < sheet.getLastRowNum(); i++) { // i=1 means start from row 2 (0-based index)
            Row row1 = sheet.getRow(i);
            Row row2 = sheet.getRow(i + 1); // next row

            if (row1 == null || row2 == null) continue;

            // Example: iterate over cells in both rows together
//            for (int j = 0; j < Math.min(row1.getLastCellNum(), row2.getLastCellNum()); j++) {
            for (int j = 2; j < sheet.getLastRowNum(); j++) {
                Cell cell1 = row1.getCell(j);
                Cell cell2 = row2.getCell(j);

                this.curr_cell = cell1;
                this.next_cell = cell2;

                this.curr_row = row1;
                this.next_row = row2;


                // Handle cell values like Python's value/value2
                String val1 = getCellValue(cell1);
                String val2 = getCellValue(cell2);






                this.curr_value = val1;
                this.next_value = val2;

                if (this.isParentPair()) {

                    this.track_parent += 1;

                    // todo: color the cells
//                    track_parent = 0;
                    parent_polymophic = 0;
                    parent_missing = 0;

                    parent1_het = 0;
                    parent2_het = 0;

                    parent1_missing = 0;
                    parent2_missing = 0;

//                    setStyleColor(red);
                    curr_cell.setCellStyle(greenstyle);
                    next_cell.setCellStyle(greenstyle);
                    // working till this point


//







                    //
                    for (int k = this.marker_col_start; k < row1.getLastCellNum(); k++) {
                        int marker_index = k - this.marker_col_start;
                        String colLetter = CellReference.convertNumToColString(k);




                        curr_value = curr_row.getCell(k).getStringCellValue();
                        next_value = next_row.getCell(k).getStringCellValue();

                        curr_cell = curr_row.getCell(k);
                        next_cell= next_row.getCell(k);

                        val1 = getCellValue(curr_cell);
                        val2 = getCellValue(next_cell);






                        String dict_value;
                        if (!val1.equals(val2) && (this.isNotMissing())) {
                          // color
//

                            this.colorAndSkipIfParentHet(colLetter);
//
                            // marker success
                        }
                        else if (this.bothMissing()) {
                            dict_value = "parent" + track_parent + colLetter;
                            dict_skip.put(dict_value, "Skip");
                            curr_cell.setCellStyle(redstyle);
                            next_cell.setCellStyle(redstyle);

                            parent1_missing += 1;
                            parent2_missing += 1;
                            parent_missing += 1;
                        }
                        else {
                            // skip
//

                            dict_value = "parent" + track_parent + colLetter;

                            dict_skip.put(dict_value, "Skip");
                            if (!homogousAndMissingSet.contains(curr_value) && !homogousAndMissingSet.contains(next_value)) {
//
                                setStyleColor(blue);

                                curr_cell.setCellStyle(bluestyle);
                                next_cell.setCellStyle(bluestyle);

                                this.parent1_het += 1;
                                this.parent2_het += 1;
                            } else {
//



                                setStyleColor(red);

//
//
//
                                curr_cell.setCellStyle(redstyle);
                                next_cell.setCellStyle(redstyle);
                            }

//
//

                        }
                    }
//
                    this.setPolymorphicHybridityValues();
                    this.colorParentGrey();
                }



            }
        }

    }


    public void f1Check() throws Exception {
        for (int i = 2; i < sheet.getLastRowNum(); i++) {
            var row1 = sheet.getRow(i);
            var row2 = sheet.getRow(i + 1); // next row

            if (row1 == null || row2 == null) continue;

            curr_row = row1;
            next_row = row2;

            if (this.f1StartandContinues()) {

                next_row.getCell(polymophic_col).setCellStyle(greystyle);
                next_row.getCell(no_parentHet_col).setCellStyle(greystyle);
                next_row.getCell(perc_parentHet_col).setCellStyle(greystyle);

                no_outcross = 0;
                missiing = 0;
                true_values = 0;


                if (curr_row.getCell(parent_col).getStringCellValue().equals("Parent") && next_row.getCell(parent_col).getStringCellValue().equals("F1")) {


                    percent_polymorph = curr_row.getCell(perc_polymorphic_col).getNumericCellValue();
                    parent_polymophic = (int) curr_row.getCell(polymophic_col).getNumericCellValue();
                    track_parent_on_F1 += 1;


                    double currentValue = curr_row.getCell(perc_parentHet_col).getNumericCellValue();
                    double previousValue = sheet.getRow(curr_row.getRowNum() - 1).getCell(perc_parentHet_col).getNumericCellValue();

                    percHet = Math.max(currentValue, previousValue);

//                        var first_parent_colx = curr_row.getCell(perc_parentHet_col).getColumnIndex();
//                        var first_parent_rowx = curr_row.getCell(perc_parentHet_col).getRowIndex();

//                        Row currentRow = row2;
//                        Row previousRow = sheet.getRow(currentRow.getRowNum() - 1);
//
//                        Cell currentCell = currentRow.getCell(percParentHetCol);
//                        Cell previousCell = previousRow.getCell(percParentHetCol);
//
//                        double currentValue = currentCell.getNumericCellValue();
//                        double previousValue = previousCell.getNumericCellValue();


//                        first_parent = f'{value[self.perc_parentHet_col].column_letter}{value[self.perc_parentHet_col].row - 1}'
//                        var first_parent= curr_row.getCell(perc_parentHet_col).getRow().;


//                        curr_row.getCell(perc_polymorphic_col);
//                        sheet.getRow(curr_row.getRowNum() - 1);


//                        self.perc_het = int(value[self.perc_parentHet_col].value) if int(value[self.perc_parentHet_col].value) > self.sheet[first_parent].value else self.sheet[first_parent].value
                }

//                for (int l = marker_start_col; l <= maxColumn; l++) {
//                    Cell cell1 = row1.getCell(l);
//                    Cell cell2 = row2.getCell(l);
//
//                    curr_cell = cell1;
//                    next_cell = cell2;

                this.colorF1Stat();
                this.determineF1hybridity();
                this.setF1Stat();
                this.setUndetermineF1();
                this.determineHybridity();
            }
            else continue;
        }
    }

    public void f1CheckBackup() throws Exception {
        for (int i = 3; i < sheet.getLastRowNum(); i++) {
            var row1 = sheet.getRow(i);
            var row2 = sheet.getRow(i + 1); // next row

            if (row1 == null || row2 == null) continue;

//            for (int j = 2; j < Math.min(row1.getLastCellNum(), row2.getLastCellNum()); j++) {
            for (int j = 4; j < sheet.getLastRowNum(); j++) {
                Cell cell1 = row1.getCell(j);
                Cell cell2 = row2.getCell(j);

                curr_cell = cell1;
                next_cell = cell2;

                curr_row = row1;
                next_row = row2;


                if (this.f1StartandContinues()) {
                    next_row.getCell(polymophic_col).setCellStyle(greystyle);
                    next_row.getCell(no_parentHet_col).setCellStyle(greystyle);
                    next_row.getCell(perc_parentHet_col).setCellStyle(greystyle);

                    no_outcross = 0;
                    missiing = 0;
                    true_values = 0;

                    if (curr_row.getCell(parent_col).getStringCellValue().equals("Parent") &&
                            next_row.getCell(parent_col).getStringCellValue().equals("F1")) {


                        percent_polymorph = curr_row.getCell(perc_polymorphic_col).getNumericCellValue();
                        parent_polymophic = (int) curr_row.getCell(polymophic_col).getNumericCellValue();
                        track_parent_on_F1 += 1;


                        double currentValue = curr_row.getCell(perc_polymorphic_col).getNumericCellValue();
                        double previousValue = sheet.getRow(curr_row.getRowNum() - 1).getCell(perc_polymorphic_col).getNumericCellValue();

                        percHet = Math.max(currentValue, previousValue);

//                        var first_parent_colx = curr_row.getCell(perc_parentHet_col).getColumnIndex();
//                        var first_parent_rowx = curr_row.getCell(perc_parentHet_col).getRowIndex();

//                        Row currentRow = row2;
//                        Row previousRow = sheet.getRow(currentRow.getRowNum() - 1);
//
//                        Cell currentCell = currentRow.getCell(percParentHetCol);
//                        Cell previousCell = previousRow.getCell(percParentHetCol);
//
//                        double currentValue = currentCell.getNumericCellValue();
//                        double previousValue = previousCell.getNumericCellValue();


//                        first_parent = f'{value[self.perc_parentHet_col].column_letter}{value[self.perc_parentHet_col].row - 1}'
//                        var first_parent= curr_row.getCell(perc_parentHet_col).getRow().;


//                        curr_row.getCell(perc_polymorphic_col);
//                        sheet.getRow(curr_row.getRowNum() - 1);


//                        self.perc_het = int(value[self.perc_parentHet_col].value) if int(value[self.perc_parentHet_col].value) > self.sheet[first_parent].value else self.sheet[first_parent].value
                    }
//                    this.determineF1hybridity();
//                    this.setF1Stat();
//                    this.colorF1Stat();
//                    this.setUndetermineF1();
//                    this.determineHybridity();
//                    return;
                }
                else {
                    continue;
                }
            }
        }
    }

    private void setUndetermineF1() {

        if (percent_polymorph <= min_perc_polymorphic) {
            next_row.getCell(hybridity_col).setCellValue("NA");
            next_row.getCell(status_col).setCellValue("Undetermine: Parent not polymorphic");
            undefined_unpolymorphic_parent += 1;
        }

        if (((next_row.getCell(perc_missing_col).getNumericCellValue() >= min_missing_percentage) ||
                        (next_row.getCell(perc_missing_col).getNumericCellValue() == 999))
                                && (isEmpty(next_row.getCell(status_col)))){

            next_row.getCell(hybridity_col).setCellValue("NA");
            next_row.getCell(status_col).setCellValue("Undetermine: missing data");
            undefined_missing_data += 1;
        }





    }

    private void color_previous(int cell) {
        var previousValue = sheet.getRow(curr_row.getRowNum() - 1).getCell(cell);
        previousValue.setCellStyle(greystyle);
    }

    private boolean isEmpty(Cell cell){
        return (cell == null) ||
                (cell.getCellType() == CellType.BLANK) ||
                (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty());
    }

    private void determineHybridity() throws Exception {

        try {

            if (isEmpty(next_row.getCell(hybridity_col))) {

                double percent_hybridity = perc_hybridity();

                if (percent_hybridity < min_perc_hybridity &&
                        isEmpty(next_row.getCell(status_col))){

                    if (percHet > max_perc_het){
                        undefined_parent_het += 1;

                        next_row.getCell(hybridity_col).setCellValue("NA");
                        next_row.getCell(status_col).setCellValue("Undetermine: Parent Heterozygous");
                    } else {
                        next_row.getCell(status_col).setCellValue("SELF");
                        FAILED += 1;
                    }
                } else {
                    next_row.getCell(status_col).setCellValue("TRUE CROSS");
                    TRUE += 1;
                }

                next_row.getCell(hybridity_col).setCellValue(percent_hybridity);
                next_row.getCell(hybridity_col).setCellStyle(greystyle);
            }


        } catch (ZeroException z) {
            curr_row.getCell(hybridity_col).setCellValue("NA");
        } catch (Exception e) {
            throw new Exception();
        }
    }

    private double perc_hybridity() {

        double another_ph;
        another_ph = (((double) true_values / (parent_polymophic - missiing)) * 100);
        return another_ph;
//        return ((new_true / (new_parent - new_missing)) * 100);
    }


    private void createStatHeaders() {

        setStyleColor(grey);

        // Create the first row
        Row headerRow = sheet.createRow(0);  // "1" in Excel = row index 0

        // Header titles
        String[] headers = {
                "#polymorphic", "%polymorphic", "#parentHet", "%parentHet",
                "#NonParentAllele", "%NonParentAllele", "#true", "#missing",
                "%missing", "%hybridity", "Status"
        };

        // Start from column C (index 2)
        int colIndex = 2;
        for (String header : headers) {
            Cell cell = headerRow.createCell(colIndex++);
            cell.setCellValue(header);
            cell.setCellStyle(greystyle);
        }
    }


    private void createPieChart(){
        int TOTAL = TRUE + FAILED + undefined_missing_data + undefined_unpolymorphic_parent + undefined_parent_het;

        var hybriditySheet = workbook.createSheet("Hybridity");
        Row header = hybriditySheet.createRow(0);
        header.createCell(0).setCellValue("Status");
        header.createCell(1).setCellValue("value");

        Object[][] data = {
                {"TRUE", TRUE},
                {"FAILED", FAILED},
                {"Undetermined: Parent not polymorphic", undefined_unpolymorphic_parent},
                {"Undetermined: Missing data", undefined_missing_data},
                {"Undetermined:Parent Heterogyte", undefined_parent_het},
                {"Total", TOTAL}
        };

        int rowIdx = 1;
        for (Object[] rowData : data) {
//        for (Object[] rowData : data) {
//            System.out.println(rowIdx);
            Row row = hybriditySheet.createRow(rowIdx++);
//            System.out.println(rowIdx);
            System.out.println(rowData[0]);
            System.out.println(rowData[1]);
            row.createCell(0).setCellValue( (String) rowData[0]);
            row.createCell(1).setCellValue( (Integer) rowData[1]);
//            row.createCell(1).setCellValue((double) rowData[1];
        }
////        hybriditySheet.

        // --- 2️⃣ Create Drawing Canvas ---
        XSSFDrawing drawing = (XSSFDrawing) hybriditySheet.createDrawingPatriarch();

        // Anchor defines position (col1, row1, col2, row2)
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);

        // --- 3️⃣ Create Chart ---
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Hybridity");
        chart.setTitleOverlay(false);

        // --- 4️⃣ Define Chart Data Source ---
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) hybriditySheet,
                new CellRangeAddress(1, data.length, 0, 0) // Category range (A2:A5)
        );
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) hybriditySheet,
                new CellRangeAddress(1, data.length, 1, 1) // Value range (B2:B5)
        );

        // --- 5️⃣ Create Pie Chart Data ---
        XDDFChartData chartData = chart.createData(ChartTypes.PIE, null, null);
        chartData.setVaryColors(true); // Enable automatic color variation
        chartData.addSeries(categories, values);

        // --- 6️⃣ Plot Chart ---
        chart.plot(chartData);

        // Optional: Add legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

    }



    private void colorF1Stat() {

        setStyleColor(grey);

        int[] cells = {
                missing_col, hybridity_col, trueF1_col, no_outcrossing_col,
                hybridity_col, status_col, perc_polymorphic_col,
                perc_missing_col, perc_outcrossing_col
        };

        for (int cell : cells) {
            next_row.getCell(cell).setCellStyle(greystyle);

//
            if (isEmpty(curr_row.getCell(cell))){
                curr_row.getCell(cell).setCellStyle(greystyle);
                this.color_previous(cell);

            }

        }

    }

    private void determineF1hybridity() {
        for (int l = marker_start_col; l <= maxColumn; l++) {
            Cell cell1 = curr_row.getCell(l);
            Cell cell2 = next_row.getCell(l);

            curr_cell = cell1;
            next_cell = cell2;


            String colLetter = CellReference.convertNumToColString(l);
            //
            //
            var testskip = "parent" + track_parent_on_F1 + colLetter;

            //            if (!(track_parent_on_F1 > 30)) {


            //            }
            //            var vartestskip = "parent" + String.valueOf(track_parent_on_F1) + curr_cell.getColumnIndex();
            //
            //
            Object dict__skip = dict_skip.get(testskip);
            String nextVal = getCellValue(next_row.getCell(l));


            if (dict__skip instanceof List<?>) {
                var list_dict__skip = (List<String>) dict_skip.get(testskip);

                if (!homogousAndMissingSet.contains(nextVal) && list_dict__skip.contains(nextVal)) {

                    next_row.getCell(l).setCellStyle(orangestyle);

                    true_values += 1;
                } else if (
                        !homogousAndMissingSet.contains(nextVal) &&
                                !list_dict__skip.contains(nextVal)) {
                    next_row.getCell(l).setCellStyle(magentastyle);
                    no_outcross += 1;

                } else if (missing_values.contains(nextVal)) {

                    missiing += 1;
                }

            } else if (dict__skip instanceof String) {
                var str_dict__skip = (String) dict_skip.get(testskip);

                if (missing_values.contains(nextVal) && !str_dict__skip.equals("Skip")) {

                    missiing += 1;
                } else if (
                        !homogous_value.contains(nextVal) &&
                                nextVal.equals(str_dict__skip) &&
                                !str_dict__skip.equals("Skip")) {

                    setStyleColor(magenta);
                    next_row.getCell(l).setCellStyle(magentastyle);
                    //                    curr_cell.setCellStyle(magentastyle);
                    no_outcross += 1;

                }


            } else if (dict__skip == null) {

            } else {

            }


        }
    }

    private void determineF1hybriditybackup(){
//
        boolean printted = false;

        if (!printted){


            printted = true;
        }


        for (int i = marker_start_col; i <= maxColumn; i++){

            String colLetter = CellReference.convertNumToColString(i);
//
//
            var testskip = "parent" + track_parent_on_F1 + colLetter;

//            if (!(track_parent_on_F1 > 30)) {


//            }
//            var vartestskip = "parent" + String.valueOf(track_parent_on_F1) + curr_cell.getColumnIndex();

            Object dict__skip = dict_skip.get(testskip);
            String nextVal = getCellValue(next_row.getCell(i));




//            if (dict__skip == null) {

//                return;
//            }


            if (dict__skip instanceof List<?>) {
                var list_dict__skip = (List<String>) dict_skip.get(testskip);
                if (list_dict__skip.contains(nextVal))

                if (!homogousAndMissingSet.contains(nextVal) && list_dict__skip.contains(nextVal)){

//                    setStyleColor(orange);

                    next_row.getCell(i).setCellStyle(orangestyle);

                    true_values += 1;
                }

            }

            else if (dict__skip instanceof String) {
                var str_dict__skip = (String) dict_skip.get(testskip);


                if (missing_values.contains(nextVal) && !str_dict__skip.equals("Skip")){

                    missiing += 1;

                }

                else if (
                        !homogous_value.contains(nextVal) &&
                                nextVal.equals(str_dict__skip) &&
                                    !str_dict__skip.equals("Skip"))
                {

                    setStyleColor(magenta);
                    next_row.getCell(i).setCellStyle(magentastyle);
//                    curr_cell.setCellStyle(magentastyle);
                    no_outcross += 1;

                }


            }  else if (dict__skip == null)   {





                continue;
            }
            else {

            }

//            if (!homogous_and_missing_value.contains(next_value) && dict__skip.contains(testskip)){
//                setStyleColor(orange);
//                true_values += 1;
//            } else {
//                if (missing_values.contains(next_value) && dict__skip.equals("Skip")) missiing += 1;
//
//                else if (){
//
//                }
//            };
        }
    }

    private void setF1Stat(){
//
//
//
//
//


        next_row.getCell(missing_col).setCellValue(missiing);
        next_row.getCell(no_outcrossing_col).setCellValue(no_outcross);
        next_row.getCell(trueF1_col).setCellValue(true_values);


        try {
            next_row.getCell(perc_missing_col).setCellValue(this.calc_perc_missing());
        } catch (Exception e){
            next_row.getCell(perc_missing_col).setCellValue("NA");
        }

        try {
            next_row.getCell(perc_outcrossing_col).setCellValue(this.calc_perc_outcross());
        } catch (Exception e){
            next_row.getCell(perc_outcrossing_col).setCellValue(0);
        }

    }

    private double calc_perc_outcross() {
        double percentage_outcross;
        try {
            percentage_outcross = (double) (no_outcross / (parent_polymophic - missiing)) * 100;
        } catch (Exception e) {
            percentage_outcross = 0;
        }
        return percentage_outcross;
    }

    private double calc_perc_missing() {
        double percentage_missing;
        try {
            percentage_missing = (double) missiing / parent_polymophic * 100;
        } catch (Exception e) {
            percentage_missing = 0;
        }
        return percentage_missing;
    }

    private boolean f1StartandContinues() {
        return ((curr_row.getCell(parent_col).getStringCellValue().equals("Parent")
                    && next_row.getCell(parent_col).getStringCellValue().equals("F1"))
                ||
                (curr_row.getCell(parent_col).getStringCellValue().equals("F1")
                        && next_row.getCell(parent_col).getStringCellValue().equals("F1")));
    }

    private void setPolymorphicHybridityValues() {
        this.set_heterozygote_columns();
        this.set_polymorphic_columns();
        this.set_missing_columns();
    }

    private boolean bothMissing() {
        var missing_values = List.of(new String[]{"Uncallable", "?"});

        return curr_value.equals(next_value) && missing_values.contains(curr_value);
    }

    private boolean isNotMissing(){
        return (!curr_value.equals("Uncallable") && !next_value.equals("Uncallable"))
                        && (!curr_value.equals("?") && !next_value.equals("?"));
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }


    private boolean isParentPair(){
        return curr_row.getCell(parent_col).getStringCellValue().equals("Parent") &&
                next_row.getCell(parent_col).getStringCellValue().equals("Parent");
    }


    private void colorAndSkipIfParentHet(String colLetter){

        String dict_value;
        if (this.heterozyte_values.contains(curr_value)){
            setStyleColor(blue);

            curr_cell.setCellStyle(bluestyle);
            next_cell.setCellStyle(bluestyle);
            dict_value = "parent" + track_parent + colLetter;
            dict_skip.put(dict_value, "Skip");
            parent1_het += 1;
        }
        else if (this.heterozyte_values.contains(next_value)) {
            setStyleColor(blue);

            curr_cell.setCellStyle(bluestyle);
            next_cell.setCellStyle(bluestyle);

            dict_value = "parent" + track_parent + colLetter;
            dict_skip.put(dict_value, "Skip");
            parent2_het += 1;
        }
        else {
            setStyleColor(green);

            curr_cell.setCellStyle(greenstyle);
            next_cell.setCellStyle(greenstyle);

            var exp3 = String.valueOf(curr_value).charAt(0) +":"+ String.valueOf(next_value).charAt(0);
            var exp4 = String.valueOf(next_value).charAt(0) +":"+ String.valueOf(curr_value).charAt(0);
            dict_value = "parent" + track_parent + colLetter;

            dict_skip.put(dict_value, List.of(exp3, exp4));
        }

        this.parent_polymophic += 1;
    }

    private void setStyleColor(XSSFColor color) {
//        for
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void set_heterozygote_columns(){

        var curr_row = curr_cell.getRow();
        var next_row = next_cell.getRow();

        curr_row.getCell(4).setCellValue(parent1_het);
        next_row.getCell(4).setCellValue(parent2_het);

//        curr_cell.setCellValue(parent1_het);
//        next_cell.setCellValue(parent2_het);

        var perc_het1 = ((float) parent1_het / no_markers) * 100;
        var perc_het2 = ((float) parent2_het / no_markers) * 100;

//
//
//

        curr_row.getCell(perc_parentHet_col).setCellValue(perc_het1);
        next_row.getCell(perc_parentHet_col).setCellValue(perc_het2);
//        curr_cell.setCellValue(perc_het1);
//        next_cell.setCellValue(perc_het2);


        // New code

//        Cell currCell4 = curr_row.getCell(4);
//        if (currCell4 == null) {
//            currCell4 = curr_row.createCell(4);
//        }
//
//        Cell nextCell4 = next_row.getCell(4);
//        if (nextCell4 == null) {
//            nextCell4 = next_row.createCell(4);
//        }
//
//        currCell4.setCellValue(parent1_het);
//        nextCell4.setCellValue(parent2_het);
//
//        float perc_het1 = ((float) parent1_het / no_markers) * 100;
//        float perc_het2 = ((float) parent2_het / no_markers) * 100;
//
//        // You might want to put these in different columns instead of overwriting the same cell
//        Cell currPercCell = curr_row.getCell(5);
//        if (currPercCell == null) {
//            currPercCell = curr_row.createCell(5);
//        }
//
//        Cell nextPercCell = next_row.getCell(5);
//        if (nextPercCell == null) {
//            nextPercCell = next_row.createCell(5);
//        }
//
//        currPercCell.setCellValue(perc_het1);
//        nextPercCell.setCellValue(perc_het2);

    }

    private void set_polymorphic_columns(){
        double perc_poly = ((double) parent_polymophic / no_markers) * 100;

        curr_row.getCell(polymophic_col).setCellValue(parent_polymophic);
        next_row.getCell(polymophic_col).setCellValue(parent_polymophic);


        curr_row.getCell(perc_polymorphic_col).setCellValue(perc_poly);
        next_row.getCell(perc_polymorphic_col).setCellValue(perc_poly);


    }

    private void set_missing_columns(){

        curr_row.getCell(missing_col).setCellValue(parent_missing);
        next_row.getCell(missing_col).setCellValue(parent_missing);


        setStyleColor(grey);

        curr_row.getCell(missing_col).setCellStyle(greystyle);
        next_row.getCell(missing_col).setCellStyle(greystyle);

        var perc_missing = (parent_missing / parent_polymophic) * 100;

        curr_row.getCell(perc_missing_col).setCellValue(perc_missing);
        next_row.getCell(perc_missing_col).setCellValue(perc_missing);
    }

    private void colorParentGrey() {

        setStyleColor(grey);

        curr_row.getCell(polymophic_col).setCellStyle(greystyle);
        next_row.getCell(polymophic_col).setCellStyle(greystyle);
        curr_row.getCell(no_parentHet_col).setCellStyle(greystyle);
        next_row.getCell(no_parentHet_col).setCellStyle(greystyle);
        curr_row.getCell(perc_parentHet_col).setCellStyle(greystyle);
        next_row.getCell(perc_parentHet_col).setCellStyle(greystyle);
        curr_row.getCell(perc_polymorphic_col).setCellStyle(greystyle);
        next_row.getCell(perc_polymorphic_col).setCellStyle(greystyle);
        curr_row.getCell(perc_missing_col).setCellStyle(greystyle);
        next_row.getCell(perc_missing_col).setCellStyle(greystyle);
    }
}
