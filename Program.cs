using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace RenameFiles
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var targetFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "Target");
            if (!Directory.Exists(targetFolderPath))
            {
                Directory.CreateDirectory(targetFolderPath);
            }
            var excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Test.xlsx");
            Console.WriteLine("注意：重命名文件放置于Target文件夹中，Excel名称修改为Test.xlsx，并与此程序放置在同一目录！");
            Console.WriteLine("注意：保证表格序号为第一列，姓名为第二列！");
            Console.WriteLine("注意：文件准备完成后按回车键继续！");
            Console.ReadKey();
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine("错误：Test.xlsx文件不存在");
                Console.ReadKey();
                return;
            }
            while (true)
            {
                Console.WriteLine("请输入结果文件夹名称：");
                var resultFolderName = Console.ReadLine();
                if (string.IsNullOrEmpty(resultFolderName))
                {
                    resultFolderName = "Result";
                }
                var resultFolderPath = Path.Combine(Directory.GetCurrentDirectory(), resultFolderName);
                if (!Directory.Exists(resultFolderPath))
                {
                    Directory.CreateDirectory(resultFolderPath);
                }
                Console.WriteLine("请输入要对比的是第几个工作表（默认为1）：");
                var sheetIndexString = Console.ReadLine();
                var sheetIndex = 0;
                if (!string.IsNullOrEmpty(sheetIndexString))
                {
                    sheetIndex = int.Parse(sheetIndexString) - 1;
                }
                RenameFilesFromExcel(excelFilePath, targetFolderPath, sheetIndex, resultFolderPath);
                Console.WriteLine("是否还有别的工作表需要筛选？（y/n）");
                var answer = Console.ReadLine();
                if (answer == "y")
                {
                    continue;
                }
                else
                {
                    break;
                }
            }
            Console.WriteLine("处理完成。");
            Console.ReadKey();
            return;
        }

        static void RenameFilesFromExcel(string excelFilePath, string targetFolderPath, int sheetIndex, string resultFolderPath)
        {
            Console.WriteLine("===============================");
            Console.WriteLine($"开始重命名文件到 {resultFolderPath}");

            // 读取Excel文件
            using (var fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new XSSFWorkbook(fs);
                var worksheet = workbook.GetSheetAt(sheetIndex);

                for (int row = 1; row <= worksheet.LastRowNum; row++)
                {
                    IRow excelRow = worksheet.GetRow(row);
                    if (excelRow == null) continue;

                    string serialNumber = excelRow.GetCell(0)?.ToString();
                    string name = excelRow.GetCell(1)?.ToString();

                    if (string.IsNullOrEmpty(name)) continue;

                    // 构建新的文件名
                    string newFileNamePrefix = serialNumber;

                    // 处理文件
                    RenameFiles(targetFolderPath, name, newFileNamePrefix, resultFolderPath);
                }
            }

            Console.WriteLine("重命名完成。");
            Console.WriteLine("===============================");
        }

        static void RenameFiles(string targetFolderPath, string name, string newFileNamePrefix, string resultFolderPath)
        {
            var files = Directory.GetFiles(targetFolderPath).Where(file => Path.GetFileName(file).Contains(name)).ToArray();
            var direcories = Directory.GetDirectories(targetFolderPath).Where(file => Path.GetFileName(file).Contains(name)).ToArray();

            if (files.Length == 0 && direcories.Length == 0)
            {
                Console.WriteLine($"未找到 {newFileNamePrefix}-{name}");
                string newFileName = $"{newFileNamePrefix}-{name}（未找到）";
                File.WriteAllText(Path.Combine(resultFolderPath, newFileName), string.Empty);
            }
            else
            {
                foreach (var oldDirPath in direcories)
                {
                    string newDirName = $"{newFileNamePrefix}-{name}";
                    string newDirPath = Path.Combine(resultFolderPath, newDirName);
                    CopyFolder(oldDirPath, newDirPath);
                }

                foreach (var oldFilePath in files)
                {
                    string suffix = Path.GetExtension(oldFilePath);
                    string newFileName = $"{newFileNamePrefix}-{name}{suffix}";
                    string newFilePath = Path.Combine(resultFolderPath, newFileName);
                    var oldFile = new FileInfo(oldFilePath);

                    oldFile.CopyTo(newFilePath, true);
                }
            }
        }
        static void CopyFolder(string sourceFolderPath, string targetFolderPath)
        {
            //如果目标路径不存在,则创建目标路径
            if (!Directory.Exists(targetFolderPath))
            {
                Directory.CreateDirectory(targetFolderPath);
            }
            //得到原文件根目录下的所有文件
            string[] files = Directory.GetFiles(sourceFolderPath);
            foreach (string file in files)
            {
                string name = Path.GetFileName(file);
                string dest = Path.Combine(targetFolderPath, name);
                File.Copy(file, dest);//复制文件
            }
            //得到原文件根目录下的所有文件夹
            string[] folders = Directory.GetDirectories(sourceFolderPath);
            foreach (string folder in folders)
            {
                string name = Path.GetFileName(folder);
                string dest = Path.Combine(targetFolderPath, name);
                CopyFolder(folder, dest);//构建目标路径,递归复制文件
            }
            return;
        }
    }
}
