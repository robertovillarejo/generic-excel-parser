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

public class ExcelRowMapParser {

    private SortedMap<Integer, String> mappingSchema;

    private Map<String, String> namedMappingSchema;

    /**
     * @param mappingSchema
     *            El 'mappingSchema' es un mapa que contiene entradas numeroColumna
     *            -> nombrePropiedad Las claves de este mapa es un subconjunto de
     *            las columnas del esquema del Workbook y los valores son los
     *            nombres de las propiedades a las que se mapearán las celdas
     */
    public ExcelRowMapParser(SortedMap<Integer, String> mappingSchema) {
        this.mappingSchema = mappingSchema;
    }

    /**
     * @param headersRow
     *            El Row que contiene los headers de la Sheet
     * @param namedMappingSchema
     *            El 'mappingSchema' es la definición del mapeo de una columna a una
     *            propiedad del objeto. Se expresa como una lista de
     *            'Header,propiedad' separados por dos puntos ':' Si el header y la
     *            propiedad tienen el mismo nombre entonces no es necesario usar
     *            'Header:propiedad', basta con escribir uno solo: 'propiedad'.
     */
    public ExcelRowMapParser(Row headersRow, Map<String, String> namedMappingSchema) {
        this.namedMappingSchema = namedMappingSchema;
        this.mappingSchema = getMappingSchema(headersRow, namedMappingSchema);
    }

    /**
     * Obtener el esquema de mapeo basado en nombres, es decir, 'Header:propiedad'
     * 
     * @return esquema de mapeo basado en nombres
     */
    public Map<String, String> getNamedMappingSchema() {
        return namedMappingSchema;
    }

    /**
     * Maps the row to a map as the defined mappingSchema
     * 
     * @param row
     * @return el mapa
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

    public static SortedMap<Integer, String> getMappingSchema(SortedMap<Integer, String> positionBasedSchema,
            Map<String, String> namedMappingSchema) {

        if (positionBasedSchema == null || namedMappingSchema == null) {
            return Collections.emptySortedMap();
        }

        TreeMap<Integer, String> mappingSchema = new TreeMap<Integer, String>();

        Set<Integer> keySet = positionBasedSchema.keySet();
        for (Integer key : keySet) {
            String value = positionBasedSchema.get(key);
            if (namedMappingSchema.containsKey(value)) {
                mappingSchema.put(key, namedMappingSchema.get(value));
            }
        }

        return mappingSchema;
    }

    public static SortedMap<Integer, String> parsePositionBasedSchema(String positionBasedSchemaExpression) {
        SortedMap<Integer, String> positionBasedSchema = new TreeMap<Integer, String>();
        String[] headers = positionBasedSchemaExpression.split(",");
        for (int i = 0; i < headers.length; i++) {
            positionBasedSchema.put(i, headers[i]);
        }
        return positionBasedSchema;
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
     * @return el esquema de mapeo basado en número de columnas
     */
    public static SortedMap<Integer, String> getMappingSchema(Row headersRow, Map<String, String> namedMappingSchema) {
        if (headersRow == null) {
            return Collections.emptySortedMap();
        }
        return ExcelRowMapParser.getMappingSchema(getPositionBasedSchema(headersRow), namedMappingSchema);
    }

    /**
     * Construye un mapa de númeroColumna -> encabezado
     * 
     * @param row
     * @return el esquema del Workbook
     */
    public static SortedMap<Integer, String> getPositionBasedSchema(Row row) {
        Iterator<Cell> cellIt = row.iterator();
        SortedMap<Integer, String> headersPositions = new TreeMap<Integer, String>();

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

    /**
     * Construye un esquema de mapeo basado en nombres a partir del
     * colonSeparatedMappingSchema
     * 
     * @param colonSeparatedMappingSchema
     * @return el esquema de mapeo basado en nombres
     */
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
