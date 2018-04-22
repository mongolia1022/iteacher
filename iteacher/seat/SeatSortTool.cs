using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using com.ccfw.Utility;
using iteacher.seat.model;
using JbkConsole;

namespace iteacher.seat
{
    /// <summary>
    /// 随机排座位工具
    /// </summary>
    public class SeatSortTool
    {
        public void Excute()
        {
            //读取学生表
            List<Student> students = null;

            try
            {
                students = new ExcelHelper().ExcelToDataTable(@".\datasource\学生名单.xlsx", "Sheet1", true)
                    .Rows.Cast<DataRow>().Select(r => new Student
                    {
                        姓名=r["姓名"].ToString(),
                        学号=ConvertHelper.ObjectToInt(r["学号"])
                    }).ToList();
            }
            catch (Exception ex)
            {
                throw new Exception("信息读取失败，请检查\"\\datasource\\学生名单.xlsx\"文件是否存在");
            }
            
            //学生缓存
            Dictionary<string, Student> studentsDic = new Dictionary<string, Student>();
            foreach (var s in students)
            {
                if (studentsDic.ContainsKey(s.姓名))
                {
                    throw new Exception("学生表存在重复姓名："+s.姓名+"。如有必要请在重复姓名后追加编号");
                }
                
                studentsDic.Add(s.姓名,s);
            }

            //读取座位表
            var seatDt=new ExcelHelper().ExcelToDataTable(@".\datasource\座位表.xlsx", "Sheet1", true);
            var seatColums = seatDt.Columns.Cast<DataColumn>().Where(c=>Regex.IsMatch(c.ColumnName,"第.+列")).ToList();
            
            //过滤人工设定的学生
            var fixedStudents = new HashSet<string>();
            foreach (DataRow row in seatDt.Rows)
            {
                foreach (var colum in seatColums)
                {
                    var studentName = row[colum.ColumnName].ToString();
                    if (!string.IsNullOrEmpty(studentName))
                    {
                        //学生表找不到姓名
                        if (!studentsDic.ContainsKey(studentName))
                        {
                            throw new Exception("座位表姓名有误：" + studentName +"不在学生名单中");
                        }
                        if (fixedStudents.Contains(studentName))
                        {
                            throw new Exception("座位表存在重复姓名：" + studentName + "。如有必要请在重复姓名后追加编号，但要与学生表一致");
                        }
                        
                        fixedStudents.Add(studentName);
                        studentsDic.Remove(studentName);
                    }
                }
            }
            
            //将过滤后的学生随机排序
            students = studentsDic.Values.OrderBy(s => Guid.NewGuid()).ToList();
            
            //写入未手工设定的座位
            foreach (DataRow row in seatDt.Rows)
            {
                
                foreach (var colum in seatColums)
                {
                    var studentName = row[colum.ColumnName].ToString();
                    //遇到单元格值为第n大组则认为座位已排完
                    if (Regex.IsMatch(studentName, "第.+大组"))
                    {
                        break;
                    }
                    
                    if (!string.IsNullOrEmpty(studentName))
                    {
                        continue;
                    }

                    var s = students.FirstOrDefault();
                    //没学生了将停止程序
                    if (s==null)
                    {
                        break;
                    }

                    row[colum.ColumnName] =s.姓名;
                    students.Remove(s);
                }
                
                if (!students.Any())
                {
                    break;
                }
            }
            
            //保存excel
            CDirectory.Create(@".\result");
            new ExcelHelper().DataTableToExcel(string.Format(@".\result\seat_sort{0}.xlsx",
                DateTime.Now.ToString("yyyy-MM-dd_HH_mm_ss")), seatDt, true).ExcuteLocal();

            Console.WriteLine("座位表生成完毕！请到.\\result目录下查找，按回车关闭应用程序");

            if (students.Any())
            {
                throw new Exception("仍有"+students.Count+"学生没有安排到座位，请核对座位表和学生名单");
            }
        }
    }
}