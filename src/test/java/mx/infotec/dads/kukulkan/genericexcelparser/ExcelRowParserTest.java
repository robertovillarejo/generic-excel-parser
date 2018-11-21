package mx.infotec.dads.kukulkan.genericexcelparser;

import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.SortedMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class ExcelRowParserTest extends TestCase {
    /**
     * Create the test case
     *
     * @param testName
     *            name of the test case
     */
    public ExcelRowParserTest(String testName) {
        super(testName);
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite() {
        return new TestSuite(ExcelRowParserTest.class);
    }

    public void testAbstractGenericExcelParser() throws IOException {
        Workbook workbook = new XSSFWorkbook(
                "/home/roberto/git/generic-excel-parser/generic-excel-parser/src/test/resources/data_example.xlsx");
        Iterator<Row> rowIt = workbook.getSheetAt(0).rowIterator();
        Row headersRow = rowIt.next();
        Map<String, String> namedMappingSchema = ExcelRowParser.parseMappingSchema(
                "evento:Partida,partida:Póliza,poliza:Fecha de transacción,fecha:Importe,total:Área,area:id");

        SortedMap<Integer, String> mappingSchema = ExcelRowParser.getMappingSchema(headersRow, namedMappingSchema);
        ExcelRowParser parser = new ExcelRowParser(mappingSchema);
        Map<String, Object> mappedInstance = parser.map(rowIt.next());
        System.out.println("Mapped Row: " + mappedInstance);
        ObjectMapper mapper = new ObjectMapper();
        Charge charge = mapper.convertValue(mappedInstance, Charge.class);
        System.out.println(charge);
        workbook.close();
    }
}
