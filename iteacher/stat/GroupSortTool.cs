using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using com.ccfw.Utility;
using CsvHelper;
using iteacher.stat.model;
using JbkConsole;

namespace iteacher.stat
{
    /// <summary>
    /// 学生分组排序工具
    /// </summary>
    public class GroupSortTool
    {
        List<string> subjects=new List<string>(); 
        public void Excute()
        {
            var groups = GroupingStudents();
            if (groups == null||!groups.Any())
            {
                Console.WriteLine("无结果数据,按回车关闭应用程序");
                return;
            }

            Dictionary<string, DataTable> dts = new Dictionary<string,DataTable>();
            foreach (var subject in subjects)
            {
                var dt = GetSubjectAvgSortTable(groups, subject);
                if(dt==null)
                    continue;

                dts.Add(subject,dt);
            }

            CDirectory.Create(@".\result");
            new ExcelHelper().DataTableToExcelWithMultSheet(string.Format(@".\result\group_sort{0}.xlsx",
                DateTime.Now.ToString("yyyy-MM-dd_HH_mm_ss")), dts, true).ExcuteLocal();

            Console.WriteLine("生成完毕！请到.\\result目录下查找，按回车关闭应用程序");
        }

        /// <summary>
        /// 获取学生分组信息
        /// </summary>
        /// <returns></returns>
        private List<StudentGroup> GetStudentGroups()
        {
            List<DataRow> rows = null;
            try
            {
                rows = new ExcelHelper().ExcelToDataTable(@".\datasource\分组名单.xlsx", "Sheet1", true)
                    .Rows.Cast<DataRow>().ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine("分组信息读取失败，请检查\"\\datasource\\分组名单.xlsx\"文件是否存在");
                return null;
            }

            var list = new List<StudentGroup>();
            foreach (var row in rows)
            {
                var groupId = ConvertHelper.ObjectToInt(row["组别"]);
                var students = row["成员"].ToString().Split(',', '，').Select(_ => new Student{姓名 = _}).ToList();
                var group = new StudentGroup
                {
                    Id = groupId,
                    Students = students
                };

                list.Add(group);
            }
            return list;
        }

        /// <summary>
        /// 获取学生各科成绩
        /// </summary>
        /// <returns></returns>
        private List<Student> GetStudentScores()
        {
            DataTable dt = null;
            try
            {
                dt = new ExcelHelper().ExcelToDataTable(@".\datasource\成绩汇总.xlsx", "Sheet1", true);
            }
            catch (Exception ex)
            {
                Console.WriteLine("分组信息读取失败，请检查\".\\datasource\\成绩汇总.xlsx\"文件是否存在");
                return null;
            }

            var students = new List<Student>();
            var studentType = typeof (Student);
            foreach (DataRow row in dt.Rows)
            {
                var student = new Student();
                foreach (var prop in studentType.GetProperties())
                {
                    if (!dt.Columns.Contains(prop.Name))
                        continue;

                    if(row[prop.Name] is DBNull)
                        continue;

                    object v = Convert.ChangeType(row[prop.Name], prop.PropertyType);

                    prop.SetValue(student, v);

                   
                }
                students.Add(student);
            }

            foreach (DataColumn colunm in dt.Columns)
            {
                if (colunm.ColumnName == "姓名" || colunm.ColumnName == "学号")
                    continue;

                subjects.Add(colunm.ColumnName);
            }

            return students;
        }

        /// <summary>
        /// 将学生分组
        /// </summary>
        /// <returns></returns>
        public List<StudentGroup> GroupingStudents()
        {
            var groups = GetStudentGroups();
            var students = GetStudentScores();

            if (groups == null || students == null)
                return null;

            foreach (var student in students)
            {
                foreach (var group in groups)
                {
                    var i = group.Students.FindIndex(_ => _.姓名 == student.姓名);
                    if (i > -1)
                        group.Students[i] = student;
                }
            }

            return groups;
        }

        private DataTable GetSubjectAvgSortTable(List<StudentGroup> groups,string subject)
        {
            var dt=new DataTable();
            dt.Columns.Add("排名");
            dt.Columns.Add("组别");
            dt.Columns.Add("姓名");
            dt.Columns.Add("分数");
            dt.Columns.Add("科目");

            var studentType = typeof (Student);
            var prop = studentType.GetProperty(subject);
            if (prop == null)
                return null;

            var group_numer = 0;
            foreach (var group in groups.OrderByDescending(_ => _.Students.Average(s=>(double)prop.GetValue(s,null))).ToList())
            {
                group_numer++;
                foreach (var student in group.Students.OrderByDescending(s => (double)prop.GetValue(s, null)).ToList())
                {
                    var row = dt.NewRow();
                    row["排名"] = group_numer;
                    row["组别"] = group.Id;
                    row["姓名"] = student.姓名;
                    row["分数"] = prop.GetValue(student);
                    row["科目"] = subject;

                    dt.Rows.Add(row);
                }
                var avgRow = dt.NewRow();
                avgRow["姓名"] = "平均分";
                avgRow["分数"] = group.Students.Average(_=>(double)prop.GetValue(_,null));
                dt.Rows.Add(avgRow);
                dt.Rows.Add(dt.NewRow());
            }

            return dt;
        }
    }
}
