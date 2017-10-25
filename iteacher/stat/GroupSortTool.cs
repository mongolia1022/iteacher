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
        public void Excute()
        {
            var groups = GroupingStudents();
            if (groups == null||!groups.Any())
            {
                Console.WriteLine("无结果数据,按回车关闭应用程序");
                return;
            }

            OutputGroupInfo(groups);
            Console.WriteLine("生成完毕！请到./result目录下查找，按回车关闭应用程序");
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
                rows = new ExcelHelper().ExcelToDataTable("./datasource/group.xlsx", "Sheet1", true)
                    .Rows.Cast<DataRow>().ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine("分组信息读取失败，请检查\"/datasource/group.xlsx\"文件是否存在");
                return null;
            }

            var list = new List<StudentGroup>();
            foreach (var row in rows)
            {
                var groupId = ConvertHelper.ObjectToInt(row["小组序号"]);
                var students = row["小组成员"].ToString().Split(',', '，').Select(_ => new Student()).ToList();
                var group = new StudentGroup
                {
                    Id = groupId,
                    Students = students
                };

                //统计
                group.Total = group.Students.Sum(_ => _.Total);
                group.Avg = group.Students.Average(_ => _.Total);

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
            List<DataRow> rows = null;
            try
            {
                rows = new ExcelHelper().ExcelToDataTable("./datasource/学生各科成绩.xlsx", "Sheet1", true)
               .Rows.Cast<DataRow>().ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine("分组信息读取失败，请检查\"/datasource/学生各科成绩.xlsx\"文件是否存在");
                return null;
            }

            var students = new List<Student>();
            foreach (var row in rows)
            {
                var student=new Student
                {
                    Name = row["姓名"].ToString(),
                    Chinese = ConvertHelper.ObjectToDouble(row["语文"]),
                    English = ConvertHelper.ObjectToDouble(row["数学"]),
                    Mathematics = ConvertHelper.ObjectToDouble(row["英语"]),
                    Politics = ConvertHelper.ObjectToDouble(row["政治"]),
                    History = ConvertHelper.ObjectToDouble(row["历史"]),
                    Geography = ConvertHelper.ObjectToDouble(row["地理"]),
                    Physics = ConvertHelper.ObjectToDouble(row["物理"]),
                    Chemistry = ConvertHelper.ObjectToDouble(row["化学"]),
                    Biology = ConvertHelper.ObjectToDouble(row["生物"]),
                };

                //统计
                student.Total = student.Chinese + student.English + student.Mathematics + student.Politics +
                                student.History + student.Geography + student.Physics + student.Chemistry +
                                student.Biology;
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
                    var i = group.Students.FindIndex(_ => _.Name == student.Name);
                    if (i > -1)
                        group.Students[i] = student;
                }
            }

            //排序
            groups=groups.OrderByDescending(_ => _.Total).ToList();

            return groups;
        }

        /// <summary>
        /// 导出分组后的成绩信息
        /// </summary>
        public void OutputGroupInfo(List<StudentGroup> groups)
        {
            FileStream fs = new FileStream(string.Format("./result_{0}.csv",DateTime.Now.ToString("yyyyMMdd")), System.IO.FileMode.Create, System.IO.FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            var csvWirter = new CsvWriter(sw);

            var number = 0;
            foreach (var group in groups)
            {
                number++;

                csvWirter.WriteRecord(new
                {
                    Number="名次",
                    GroupId="小组序号",
                    Name="学生姓名",
                    Chinese = "语文",
                    Mathematics = "数学",
                    English = "英语",
                    Politics = "政治",
                    History = "历史",
                    Geography = "地理",
                    Physics = "物理",
                    Chemistry = "化学",
                    Biology = "生物",
                });
                foreach (var student in group.Students)
                {
                    csvWirter.WriteRecord(new
                    {
                        Number = "第"+number+"名",
                        GroupId = group.Id+"组",
                        student.Name,
                        student.Chinese,
                        student.Mathematics,
                        student.English,
                        student.Politics,
                        student.History,
                        student.Geography,
                        student.Physics,
                        student.Chemistry,
                        student.Biology,
                    });
                }
                csvWirter.WriteRecord(new
                {
                    Number = "",
                    GroupId = "",
                    Name = "",
                    Chinese = "",
                    Mathematics = "",
                    English = "",
                    Politics = "",
                    History = "",
                    Geography = "",
                    Physics = "",
                    Chemistry = "",
                    Biology = "",
                });
            }
        }
    }
}
