package org.example

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream

class Student(
    var name:String,
    var controlWorksList:MutableList<Int> = mutableListOf(),
    var labsList:MutableList<Int> = mutableListOf(),
    var projectList:MutableList<Int> = mutableListOf(),
    var attendsList:MutableList<Int> = mutableListOf(),
)

fun main() {
    val groupMap = getExcel()
    println(groupMap)
}

fun getExcel():Map<String,List<Student>>{
    var resMap = mutableMapOf<String,MutableList<Student>>()

    val file  = FileInputStream(File("C:/Users/minen/Desktop/ФиЛП_2024.xlsx"))
    val workbook = XSSFWorkbook(file)

    val sheet36_1 = workbook.getSheet("36_1")
    val student36_1List = getStudentlist(17,sheet36_1,0)
    resMap.put("36_1", student36_1List)

    val sheet36_2 = workbook.getSheet("36_2")
    val student36_2List = getStudentlist(16,sheet36_2,0)
    resMap.put("36_2", student36_2List)

    val sheet39 = workbook.getSheet("39")
    val student39List = getStudentlist(16,sheet39,2)
    resMap.put("39", student39List)

    return resMap

}

fun getStudentlist(studentNum:Int, sheet:Sheet, step:Int):MutableList<Student>{

    var resList = mutableListOf<Student>()

    for(rowNum in 2..(studentNum+1)) {
        val curRow = sheet.getRow(rowNum)
        val stName = curRow.getCell(1).stringCellValue
        val st:Student = Student(name = stName)

        for (labCell in 28..34){
            val curCell = curRow.getCell(labCell+step)
            if(curCell.cellType == CellType.STRING){
                if (curCell.stringCellValue=="д") st.labsList.add(2)
            }
            if (curCell.cellType == CellType.NUMERIC){
                st.labsList.add(curCell.numericCellValue.toInt())
            }
            if (curCell.cellType == CellType.BLANK){
                st.labsList.add(0)
            }
        }

        for (attCell in 4..15){
            val curCell = curRow.getCell(attCell)
            if((curCell.cellType == CellType.STRING)){
                if (curCell.stringCellValue=="+") st.attendsList.add(1)
            }
            if (curCell.cellType == CellType.BLANK){
                st.attendsList.add(0)
            }
        }

        for (conWorkCell in 40..42){
            val curCell = curRow.getCell(conWorkCell+step)
            if(curCell.cellType == CellType.STRING){
                if (curCell.stringCellValue=="д") st.controlWorksList.add(2)
                else st.controlWorksList.add(0)
            }
            if (curCell.cellType == CellType.NUMERIC){
                st.controlWorksList.add(curCell.numericCellValue.toInt())
            }
        }

        for (exProject in 35..39){
            val curCell = curRow.getCell(exProject+step)
            if(curCell.cellType == CellType.STRING){
                if (curCell.stringCellValue=="д") st.projectList.add(2)
            }
            if (curCell.cellType == CellType.NUMERIC){
                st.projectList.add(curCell.numericCellValue.toInt())
            }
            if (curCell.cellType == CellType.BLANK){
                st.projectList.add(0)
            }
        }
        resList.add(st)
    }
    return resList
}