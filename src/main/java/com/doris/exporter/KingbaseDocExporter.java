package com.doris.exporter;

import org.apache.poi.xwpf.usermodel.*;
import java.sql.*;
import java.io.*;
import java.math.BigInteger;

public class KingbaseDocExporter {
    private static final String TABLE_QUERY = "SELECT tablename AS TABLE_NAME, description AS TABLE_COMMENT " +
            "FROM pg_tables t LEFT JOIN pg_description d ON d.objoid = t.tablename::regclass " +
            "WHERE schemaname = ?";

    private static final String COLUMN_QUERY = "SELECT a.attname AS COLUMN_NAME, pg_catalog.format_type(a.atttypid, a.atttypmod) AS COLUMN_TYPE, " +
            "d.description AS COLUMN_COMMENT, a.attnum AS ORDINAL_POSITION, " +
            "CASE WHEN a.atttypmod > 0 THEN a.atttypmod-4 ELSE NULL END AS CHARACTER_MAXIMUM_LENGTH, " +
            "CASE WHEN a.attnotnull THEN 'PRI' ELSE '' END AS COLUMN_KEY " +
            "FROM pg_attribute a LEFT JOIN pg_description d ON d.objoid = a.attrelid AND d.objsubid = a.attnum " +
            "WHERE a.attrelid = ?::regclass AND a.attnum > 0 AND NOT a.attisdropped";

    public static void main(String[] args) {
        String jdbcUrl = "jdbc:kingbase8://host?currentSchema=currentSchema";
        String username = "";
        String password = "";
        String schema = "";
        String outputFile = "output.docx";

        try {
            // Load Kingbase JDBC driver
            try {
                Driver driver = (Driver)Class.forName("com.kingbase8.Driver").newInstance();
                DriverManager.registerDriver(driver);
                System.out.println("Kingbase JDBC driver loaded and registered successfully");
            } catch (Exception e) {
                System.err.println("Failed to load Kingbase JDBC driver: " + e.getMessage());
                System.err.println("Please ensure:");
                System.err.println("1. kingbase8-8.6.0.jar is in src/main/resources/lib/ directory");
                System.err.println("2. The jar file is included in classpath");
                System.err.println("Current classpath: " + System.getProperty("java.class.path"));
                throw new RuntimeException("Failed to load JDBC driver", e);
            }
            System.out.println("Attempting to connect to database with URL: " + jdbcUrl);
            Connection conn;
            try {
                conn = DriverManager.getConnection(jdbcUrl, username, password);
                System.out.println("Database connection established successfully");
            } catch (SQLException e) {
                System.err.println("Failed to connect to database: " + e.getMessage());
                System.err.println("Please verify:");
                System.err.println("1. Database URL format is correct");
                System.err.println("2. Database is running and accessible");
                System.err.println("3. Username and password are correct");
                throw e;
            }
            XWPFDocument doc = new XWPFDocument();
            processSchema(conn, schema, doc);
            try (FileOutputStream out = new FileOutputStream(outputFile)) {
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
        String chineseName = tableComment;
        String tablePurpose = tableComment;
        doc.createParagraph().createRun().setText("表中文名称：" + chineseName);
        doc.createParagraph().createRun().setText("表英文名称：" + tableName);
        doc.createParagraph().createRun().setText("表用途：" + tablePurpose);

        XWPFTable table = doc.createTable();
        table.setWidth("100%");

        XWPFTableRow headerRow = table.getRow(0);
        for (int i = 0; i < colWidths.length; i++) {
            if (i >= headerRow.getTableCells().size()) {
                headerRow.createCell();
            }
            XWPFTableCell cell = headerRow.getCell(i);
            cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(colWidths[i]));
        }

        addTableHeaderCells(headerRow, new String[] {
                "表英文名", "字段英文名", "字段中文解释",
                "字段数据类型", "字段序号", "字段长度",
                "约束条件主键", "是否代码", "备注"
        });
    }

    private static void processColumns(Connection conn, String schema, String tableName, XWPFDocument doc)
            throws SQLException {
        try (PreparedStatement columnStmt = conn.prepareStatement(COLUMN_QUERY)) {
            columnStmt.setString(1, tableName);
            ResultSet columns = columnStmt.executeQuery();

            XWPFTable table = doc.getTables().get(doc.getTables().size() - 1);
            while (columns.next()) {
                XWPFTableRow row = table.createRow();
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
        addCell(row, row.getCell(cellIndex++), "");
        addCell(row, row.getCell(cellIndex++), "");
    }

    private static void addCell(XWPFTableRow row, XWPFTableCell cell, String value) {
        if (cell == null) {
            cell = row.createCell();
        }
        cell.setText(value);
        if (row.getTableCells().size() < colWidths.length) {
            row.createCell();
        }
    }
}