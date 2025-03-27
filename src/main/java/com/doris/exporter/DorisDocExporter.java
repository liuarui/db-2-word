package com.doris.exporter;

import org.apache.poi.xwpf.usermodel.*;
import java.sql.*;
import java.io.*;
import java.math.BigInteger;

public class DorisDocExporter {
    private static final String TABLE_QUERY = "SELECT TABLE_NAME, TABLE_COMMENT " +
            "FROM information_schema.tables " +
            "WHERE TABLE_SCHEMA = ?";

    private static final String COLUMN_QUERY = "SELECT COLUMN_NAME, COLUMN_TYPE, COLUMN_COMMENT, " +
            "ORDINAL_POSITION, CHARACTER_MAXIMUM_LENGTH, COLUMN_KEY " +
            "FROM information_schema.columns " +
            "WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?";

    public static void main(String[] args) {
        if (args.length < 4) {
            System.out.println("Usage: java DorisDocExporter <jdbc_url> <user> <password> <schema> <output.docx>");
            return;
        }

        try (Connection conn = DriverManager.getConnection(args[0], args[1], args[2])) {
            XWPFDocument doc = new XWPFDocument();
            processSchema(conn, args[3], doc);
            try (FileOutputStream out = new FileOutputStream(args[4])) {
                doc.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void processSchema(Connection conn, String schema, XWPFDocument doc) throws SQLException {
        try (PreparedStatement tableStmt = conn.prepareStatement(TABLE_QUERY)) {
            tableStmt.setString(1, schema);
            ResultSet tables = tableStmt.executeQuery();

            while (tables.next()) {
                String tableName = tables.getString("TABLE_NAME");
                String tableComment = tables.getString("TABLE_COMMENT");
                addTableSection(doc, tableName, tableComment);
                processColumns(conn, schema, tableName, doc);
            }
        }
    }

    private static final int[] colWidths = { 1000, 2500, 3500, 1500, 1000, 1500, 1500, 1500, 2500 };

    private static void addTableSection(XWPFDocument doc, String tableName, String tableComment) {
        // 解析表注释（格式：表中文名称：xxx|表用途：xxx）
        String chineseName = tableComment;
        String tablePurpose = tableComment;
        // 添加表基本信息
        doc.createParagraph().createRun().setText("表中文名称：" + chineseName);
        doc.createParagraph().createRun().setText("表英文名称：" + tableName);
        doc.createParagraph().createRun().setText("表用途：" + tablePurpose);

        // 创建表格
        XWPFTable table = doc.createTable();
        table.setWidth("100%");

        // 创建表格标题行
        XWPFTableRow headerRow = table.getRow(0);
        // 设置列宽（单位：二十分之一英寸）
        // 创建完整列结构
        for (int i = 0; i < colWidths.length; i++) {
            if (i >= headerRow.getTableCells().size()) {
                headerRow.createCell();
            }
            XWPFTableCell cell = headerRow.getCell(i);
            cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(colWidths[i]));
        }
        // 修正：这里调用方法添加表头单元格
        // 调用添加表头单元格的方法
        addTableHeaderCells(headerRow, new String[] {
                "表英文名", "字段英文名", "字段中文解释",
                "字段数据类型", "字段序号", "字段长度",
                "约束条件主键", "是否代码", "备注"
        });
    }

    private static void processColumns(Connection conn, String schema, String tableName, XWPFDocument doc)
            throws SQLException {
        try (PreparedStatement columnStmt = conn.prepareStatement(COLUMN_QUERY)) {
            columnStmt.setString(1, schema);
            columnStmt.setString(2, tableName);
            ResultSet columns = columnStmt.executeQuery();

            XWPFTable table = doc.getTables().get(doc.getTables().size() - 1);
            while (columns.next()) {
                XWPFTableRow row = table.createRow();
                // 确保列宽设置一致
                for (int i = 0; i < colWidths.length; i++) {
                    if (i >= row.getTableCells().size()) {
                        row.createCell();
                    }
                    XWPFTableCell cell = row.getCell(i);
                    cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(colWidths[i]));
                }
                addColumnCells(tableName, row, columns);
            }
        }
    }

    private static void addTableHeaderCells(XWPFTableRow row, String[] headers) {
        for (int i = 0; i < headers.length; i++) {
            XWPFTableCell cell = i < row.getTableCells().size() ? row.getCell(i) : row.createCell();
            XWPFRun run = cell.addParagraph().createRun();
            run.setText(headers[i]);
            run.setBold(true);
        }
    }

    private static void addColumnCells(String tableName,XWPFTableRow row, ResultSet column) throws SQLException {
        int cellIndex = 0;
        addCell(row, row.getCell(cellIndex++), tableName);
        addCell(row, row.getCell(cellIndex++), column.getString("COLUMN_NAME"));
        addCell(row, row.getCell(cellIndex++), column.getString("COLUMN_COMMENT"));
        addCell(row, row.getCell(cellIndex++), column.getString("COLUMN_TYPE"));
        addCell(row, row.getCell(cellIndex++), column.getString("ORDINAL_POSITION"));
        addCell(row, row.getCell(cellIndex++), column.getString("CHARACTER_MAXIMUM_LENGTH"));
        addCell(row, row.getCell(cellIndex++), column.getString("COLUMN_KEY").equals("PRI") ? "PK" : "");
        addCell(row, row.getCell(cellIndex++), ""); // 是否代码
        addCell(row, row.getCell(cellIndex++), ""); // 备注
    }

    private static void addCell(XWPFTableRow row, XWPFTableCell cell, String value) {
        if (cell == null) {
            cell = row.createCell();
        }
        cell.setText(value);
        // 保持列数一致
        if (row.getTableCells().size() < colWidths.length) {
            row.createCell();
        }
    }
}