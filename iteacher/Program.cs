using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iteacher.seat;
using iteacher.stat;

namespace iteacher
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("请输入命令启动程序：");
            Console.WriteLine("1:分组成绩排序,2:座位表生成");
            while (true)
            {
                string command = Console.ReadLine();
                if (command == "1")
                {
                    new GroupSortTool().Excute();
                    break;
                }

                if (command == "2")
                {
                    new SeatSortTool().Excute();
                    break;
                }

                Console.WriteLine("未知命令:" + command + ",请重新输入");
            }

            Console.ReadLine();
        }
    }
}
