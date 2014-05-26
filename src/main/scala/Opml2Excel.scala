import scala.xml.{NodeSeq, XML}

/**
 * Created by markmo on 26/05/2014.
 */
object Opml2Excel extends App {

  case class Row(name: String, note: String, children: Seq[Row], level: Int) {
    override def toString = {
      " " * level * 4 + name + "\n" + children.mkString("")
    }
  }

  case class Column(name: String, values: List[ColumnValue])

  case class ColumnValue(rowNum: Int, value: String)

  val columns = collection.mutable.Map[String, Column]()

  def parseOpml(s: NodeSeq, rowNum: Int, level: Int): Seq[Row] = {
    if (s.isEmpty) {
      return Nil
    }
    s.map { n =>
      n.attributes foreach { attr =>
        if (columns.contains(attr.key)) {
          columns(attr.key).values :+ attr.value.text
        } else {
          columns(attr.key) = Column(attr.key, List(ColumnValue(rowNum + 1, attr.value.text)))
        }
      }
      val note = (n \ "_note").text
      val children = parseOpml(n \ "outline", rowNum + 1, level + 1)
      Row((n \ "@text").text, note, children, level + 1)
    }
  }

  println("Converting OPML file to Excel")
  //println(args(0))
  //val xml = XML.loadFile(args(0))
  val xml = XML.loadFile("/Users/markmo/Documents/ANZ Systems Integration.opml")
  val title = (xml \ "head" \ "title").text
  println(title)
  //println(xml)
  val rows = parseOpml(xml \ "body", 0, 0)
  rows foreach { row =>
    println(row)
  }

}
