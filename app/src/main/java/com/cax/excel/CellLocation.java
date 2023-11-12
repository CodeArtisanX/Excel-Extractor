/*
 * CellLocation.java
 *
 * Description:
 * This Java class represents a cell location in an Excel sheet, with information about the row and column indices.
 *
 * Author: CodeArtisanX
 * Date: November 11, 2023
 */
package com.cax.excel;

/**
 * CellLocation class representing a cell location in an Excel sheet.
 */
public class CellLocation {
    private final int rowIndex;
    private final int colIndex;

    /**
     * Constructs a CellLocation with the specified row and column indices.
     *
     * @param rowIndex The index of the row.
     * @param colIndex The index of the column.
     */
    public CellLocation(int rowIndex, int colIndex) {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
    }

    /**
     * Gets the row index of the cell location.
     *
     * @return The row index.
     */
    public int getRowIndex() {
        return rowIndex;
    }

    /**
     * Gets the column index of the cell location.
     *
     * @return The column index.
     */
    public int getColIndex() {
        return colIndex;
    }
}

