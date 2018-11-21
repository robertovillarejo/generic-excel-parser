/*
 *  
 * The MIT License (MIT)
 * Copyright (c) 2018 Roberto Villarejo Martínez
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

package mx.infotec.dads.kukulkan.genericexcelparser;

import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class ExcelRowParser {

    private SortedMap<Integer, String> mappingSchema;

    private Map<String, String> namedMappingSchema;

    public ExcelRowParser(SortedMap<Integer, String> mappingSchema) {
        this.mappingSchema = mappingSchema;
    }

    public ExcelRowParser(Row headersRow, Map<String, String> namedMappingSchema) {
        this.namedMappingSchema = namedMappingSchema;
        this.mappingSchema = getMappingSchema(headersRow, namedMappingSchema);
    }

    public Map<String, String> getNamedMappingSchema() {
        return namedMappingSchema;
    }

    /**
     * Maps the row to a map as the defined mappingSchema
     * 
     * @param row
     * @return the map
     */
    public Map<String, Object> map(Row row) {
        Map<String, Object> mappedInstance = new HashMap<String, Object>();
        for (Entry<Integer, String> entry : mappingSchema.entrySet()) {
            Object value;
            try {
                if (CellType.NUMERIC.equals(row.getCell(entry.getKey()).getCellTypeEnum())) {
                    value = row.getCell(entry.getKey()).getNumericCellValue();
                } else if (CellType.FORMULA.equals(row.getCell(entry.getKey()).getCellTypeEnum())) {
                    value = row.getCell(entry.getKey()).getCellFormula();
                } else {
                    value = row.getCell(entry.getKey()).getStringCellValue();
                }
            } catch (Exception ex) {
                value = null;
            }
            mappedInstance.put(entry.getValue(), value);
        }
        return mappedInstance;
    }

    /**
     * Generates the mapping schema based on column positions from a headersRow and
     * namedMappingSchema
     * 
     * @param headersRow
     *            the row with headers (usually the first header)
     * @param namedMappingSchema
     *            the map with 'Header Name' (from Excel Sheet) to 'Property Name'
     *            for mapping
     * @return
     */
    public static SortedMap<Integer, String> getMappingSchema(Row headersRow, Map<String, String> namedMappingSchema) {
        TreeMap<Integer, String> mappingSchema = new TreeMap<Integer, String>();

        if (headersRow == null) {
            return Collections.emptySortedMap();
        }

        Map<Integer, String> headersPositions = getHeadersPositions(headersRow);
        Set<Integer> keySet = headersPositions.keySet();
        for (Integer key : keySet) {
            String value = headersPositions.get(key);
            if (namedMappingSchema.containsKey(value)) {
                mappingSchema.put(key, namedMappingSchema.get(value));
            }
        }

        return mappingSchema;
    }

    private static Map<Integer, String> getHeadersPositions(Row row) {
        Iterator<Cell> cellIt = row.iterator();
        Map<Integer, String> headersPositions = new HashMap<Integer, String>();

        while (cellIt.hasNext()) {
            Cell cell = cellIt.next();
            String headerString;
            try {
                headerString = cell.getStringCellValue();
            } catch (Exception ex) {
                headerString = "";
            }
            headersPositions.put(cell.getColumnIndex(), headerString);
        }
        return headersPositions;
    }

    public static Map<String, String> parseMappingSchema(String colonSeparatedMappingSchema) {
        Map<String, String> mappingSchema = new HashMap<String, String>();
        if (colonSeparatedMappingSchema == null || "".equals(colonSeparatedMappingSchema)) {
            return Collections.emptyMap();
        }
        String[] mappingProperties = colonSeparatedMappingSchema.split(":");
        for (String mappingProperty : mappingProperties) {
            String[] headerProperty = mappingProperty.split(",");
            if (headerProperty.length >= 2) {
                mappingSchema.put(headerProperty[0], headerProperty[1]);
            } else if (headerProperty.length == 1) {
                mappingSchema.put(headerProperty[0], headerProperty[0]);
            }
        }
        return mappingSchema;
    }

}
