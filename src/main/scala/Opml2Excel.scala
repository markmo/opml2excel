import java.io.FileOutputStream
import org.apache.poi.hssf.usermodel.{HSSFSheet, HSSFWorkbook}
import scala.xml.{NodeSeq, XML}

/**
 * Created by markmo on 26/05/2014.
 */
object Opml2Excel extends App {

  case class Row(name: String, note: String, children: Seq[Row], rowNum: Int, level: Int) {
    override def toString = {
      rowNum + " " * level * 4 + name + "\n" + children.mkString("")
    }
  }

  case class Column(name: String, values: List[ColumnValue]) {
    def getValue(rowNum: Int) = {
      val a = values.filter(_.rowNum == rowNum)
      if (a.isEmpty) {
        null
      } else {
        a.head.value
      }
    }
  }

  case class ColumnValue(rowNum: Int, value: String)

  val columns = collection.mutable.Map[String, Column]()

  def parseOpml(s: NodeSeq, rowNum: Int, level: Int): (Seq[Row], Int) = {
    if (s.isEmpty) {
      return (Nil, 0)
    }
    var i = rowNum
    val rows = s.map { n =>
      n.attributes foreach { attr =>
        if (attr.key != "text") {
          if (columns.contains(attr.key)) {
            columns(attr.key).values :+ attr.value.text
          } else {
            columns(attr.key) = Column(attr.key, List(ColumnValue(i, attr.value.text)))
          }
        }
      }
      val note = (n \ "_note").text
      val (children, j) = parseOpml(n \ "outline", i + 1, level + 1)
      val row = Row((n \ "@text").text, note, children, i, level)
      i += j + 1
      row
    }
    (rows, i - rowNum)
  }

  var i = 1

  def writeRows(sheet: HSSFSheet, rows: Seq[Row]): Unit = {
    rows foreach { r =>
      if (r.rowNum > 0) {
        val row = sheet.createRow(r.rowNum)
        val nameCell = row.createCell(0)
        nameCell.setCellValue(" " * r.level * 4 + r.name)
        i = 1
        columns.values foreach { c =>
          val cell = row.createCell(i)
          cell.setCellValue(c.getValue(r.rowNum))
          i += 1
        }
      }
      if (!r.children.isEmpty) {
        writeRows(sheet, r.children)
        sheet.groupRow(r.rowNum + 1, r.rowNum + r.children.length)
      }
    }
  }

  def writeExcel(title: String, rows: Seq[Row]) = {
    val wb = new HSSFWorkbook()
    val sheet = wb.createSheet(title)
    sheet.setRowSumsBelow(false)
    val headerRow = sheet.createRow(0)
    val nameCell = headerRow.createCell(0)
    nameCell.setCellValue("Title")
    columns.keys foreach { colName =>
      val cell = headerRow.createCell(i)
      cell.setCellValue(colName)
      i += 1
    }
    writeRows(sheet, rows)
    val out = new FileOutputStream("test.xls")
    wb.write(out)
    out.close()
  }

  println("Converting OPML file to Excel")
  //println(args(0))
  //val xml = XML.loadFile(args(0))
  val xml = XML.loadFile("/Users/markmo/Documents/ANZ Systems Integration.opml")
  val title = (xml \ "head" \ "title").text
  println(title)
  //println(xml)
  val (rows, k) = parseOpml(xml \ "body", 0, -1)
  rows foreach { row =>
    println(row)
  }
  writeExcel(title, rows)

}
