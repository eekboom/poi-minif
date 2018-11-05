import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import spock.lang.Specification
import spock.lang.Unroll


class PoiSpec extends Specification {

    @Unroll
    def 'test array formula: #value -> #expectedValue'() {
        given:
        Workbook workbook = WorkbookFactory.create(Thread.currentThread().getContextClassLoader().getResourceAsStream("workbook.xlsx"))
        Sheet sheet = workbook.getSheet("Sheet1")
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Row inputRow = sheet.getRow(0)
        Cell inputCell = inputRow.getCell(2)

        Row outputRow = sheet.getRow(2)
        Cell outputCell = outputRow.getCell(2)

        when:
        inputCell.setCellValue(value)
        formulaEvaluator.notifyUpdateCell(inputCell)

        then:
        formulaEvaluator.evaluate(outputCell).getNumberValue() == expectedValue

        where:
        value || expectedValue
        3     || 5
        5     || 5
        6     || 10
        27    || 30

    }
}
