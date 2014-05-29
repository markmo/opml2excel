import java.io.FileOutputStream
import org.apache.poi.hssf.usermodel.{HSSFSheet, HSSFWorkbook}
import org.apache.poi.ss.usermodel.{Font, CellStyle}
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

    def isNumeric =
      values.forall((x) => Opml2Excel.isNumeric(x.value))

  }

  def isNumeric(x: String) = {
    if (x == null || x.isEmpty) {
      false
    } else {
      x.matches(s"""[+-]?((\\d+(e\\d+)?[lL]?)|(((\\d+(\\.\\d*)?)|(\\.\\d+))(e\\d+)?[fF]?))""")
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
            columns(attr.key) = Column(attr.key, columns(attr.key).values :+ ColumnValue(i, attr.value.text))
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

  val letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  var i = 1

  def writeRows(sheet: HSSFSheet, rows: Seq[Row]): Int = {
    var j = 0
    rows foreach { r =>
      if (r.rowNum > 0) {
        val row = sheet.createRow(r.rowNum)
        val titleStyle = getStyle(sheet.getWorkbook, r.level)
        val colStyle = sheet.getWorkbook.createCellStyle()
        colStyle.setWrapText(true)
        colStyle.setVerticalAlignment(CellStyle.VERTICAL_TOP)
        val titleCell = row.createCell(0)
        titleCell.setCellValue(r.name)
        titleCell.setCellStyle(titleStyle)
        i = 1
        columns.values foreach { c =>
          val cell = row.createCell(i)
          val numberChildren = r.children.length
          if (false && numberChildren > 0 && c.isNumeric) {
            val colRef = letters(i - 1)
            val formula = "SUM(" + colRef + r.rowNum + ":" + colRef + r.rowNum + numberChildren + ")"
            cell.setCellFormula(formula)
          } else {
            if (c.name == "_note") {
              cell.setCellStyle(colStyle)
            }
            val v = c.getValue(r.rowNum)
            if (isNumeric(v)) {
              cell.setCellValue(v.toDouble)
            } else {
              cell.setCellValue(v)
            }
          }
          i += 1
        }
      }
      if (!r.children.isEmpty) {
        val k = writeRows(sheet, r.children)
        if (r.rowNum > 0) {
          sheet.groupRow(r.rowNum + 1, r.rowNum + k)
        }
        j += k + 1
      } else {
        j += 1
      }
    }
    j
  }

  var styles = collection.mutable.Map[Int, CellStyle]()

  def getStyle(wb: HSSFWorkbook, indent: Int) = {
    if (styles.contains(indent)) {
      styles(indent)
    } else {
      val style = wb.createCellStyle()
      style.setWrapText(true)
      style.setVerticalAlignment(CellStyle.VERTICAL_TOP)
      style.setIndention(indent.toShort)
      styles(indent) = style
      style
    }
  }

  def writeExcel(title: String, rows: Seq[Row], filename: String) = {
    val wb = new HSSFWorkbook()
    val sheet = wb.createSheet(title)
    sheet.setRowSumsBelow(false)
    val headerStyle = wb.createCellStyle()
    val headerFont = wb.createFont()
    headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD)
    headerStyle.setFont(headerFont)
    val headerRow = sheet.createRow(0)
    val nameCell = headerRow.createCell(0)
    nameCell.setCellValue("Title")
    nameCell.setCellStyle(headerStyle)
    columns.keys foreach { colName =>
      val cell = headerRow.createCell(i)
      if (colName == "_note") {
        cell.setCellValue("Notes")
      } else {
        cell.setCellValue(colName)
      }
      cell.setCellStyle(headerStyle)
      i += 1
    }
    writeRows(sheet, rows)
    sheet.setColumnWidth(0, 60*256)
    sheet.setColumnWidth(1, 60*256)
    val out = new FileOutputStream(filename)
    wb.write(out)
    out.close()
  }

  println("Converting OPML file to Excel")
  //println(args(0))
  val infilename = args(0)
  val xml = XML.loadFile(infilename)
  val outfilename = infilename.slice(0, infilename.length - 4) + "xls"
  //val xml = XML.loadFile("/Users/markmo/Documents/ANZ Systems Integration.opml")
  val title = (xml \ "head" \ "title").text
  println(title)
  //println(xml)
  val (rows, k) = parseOpml(xml \ "body", 0, -1)
  /*
  rows foreach { row =>
    println(row)
  }
  columns foreach {
    case (name, column) =>
      column.values foreach { v =>
        println(v.rowNum + ": " + v.value)
      }
  }*/
  writeExcel(title, rows, outfilename)

}
