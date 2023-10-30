using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using ClosedXML.Excel;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Bibliography;

namespace TestTaskAkelon
{
    class Program
    {

        static void Main(string[] args)
        {
            int resultLoad = args.Length > 0 ? LoadInfoFromFile(args[0]) : LoadInfoFromFile("excel.xlsx");
            if (resultLoad != 1)
            {
                Console.ReadKey();
                return;
            }
            while (true)
                ChooseAction();
        }

        static DataSet dsMain = new DataSet();
        static string[] requiredTables = { "Товары", "Клиенты", "Заявки" };
        static FileExcel fileExcel;

        /// <summary>
        /// Выбор действия
        /// </summary>
        static void ChooseAction()
        {
            Console.WriteLine("Выбирите необходимое действие и введите цифру\n1 - Вывод информации по товару\n2 - Изменить контактное лицо организации" +
                "\n3 - Вывести \"Золотого\" клиента\n4 - Загрузить новый файл\n5 - Выход");
            string variant = Console.ReadLine();
            switch (variant)
            {
                case "1":
                    ViewInfoOfClientByProduct();
                    break;
                case "2":
                    ChangeContactPerson();
                    break;
                case "3":
                    FindGoldenClient();
                    break;
                case "4":
                    LoadInfoFromFile();
                    break;
                case "5":
                    Environment.Exit(0);
                    break;
                default:
                    Console.WriteLine("Указанная команда не распознаан, введите цифру варианта действия");
                    break;
            }
        }
       
        /// <summary>
        /// Вывод информации о клиентах, заказавших товар с указанием информации по количеству товара, цене и дате заказа
        /// </summary>
        static void ViewInfoOfClientByProduct()
        {
            if (!CheckRequiredTables() || PrintInfoFromTable("Товары") == -1)
            {
                return;
            }

            string productName = "";
            while (productName == String.Empty)
            {
                Console.WriteLine("Введите название товара для вывода информации");
                productName = Console.ReadLine();
            }

            CheckRequiredTables();

            var result = from goods in dsMain.Tables["Товары"].AsEnumerable()
                         join req in dsMain.Tables["Заявки"].AsEnumerable() on goods.Field<int>("Код товара") equals req.Field<int>("Код товара")
                         join client in dsMain.Tables["Клиенты"].AsEnumerable() on req.Field<int>("Код клиента") equals client.Field<int>("Код клиента")
                         where goods.Field<string>("Наименование") == productName
                         select new
                         {
                             name = client.Field<string>("Наименование организации"),
                             adres = client.Field<string>("Адрес"),
                             contact = client.Field<string>("Контактное лицо (ФИО)"),
                             requiredId = req.Field<int>("Номер заявки"),
                             goodCount = req.Field<int>("Требуемое количество"),
                             requiredDate = req.Field<DateTime>("Дата размещения"),
                             goodPrice = goods.Field<int>("Цена товара за единицу"),
                             finalPrice = goods.Field<int>("Цена товара за единицу") * req.Field<int>("Требуемое количество")
                         };
            if (result.Count() > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var info in result)
                {
                    sb.AppendLine(String.Format("Наименование организации - {0,-50}\nАдрес - {1,-50}\nКонтактное лицо (ФИО) - {2,-50}\nНомер заявки - {3,-50}\n" +
                        "Требуемое количество - {4,-50}\nДата размещения - {5,-50}\nЦена товара за единицу - {6,-50}\nВсего- {7,-50}",
                        info.name,
                        info.adres,
                        info.contact,
                        info.requiredId,
                        info.goodCount,
                        info.requiredDate,
                        info.goodPrice,
                        info.finalPrice));
                }
                Console.WriteLine(sb.ToString());
            }
            else
            {
                Console.WriteLine("Товар еще не заказывали");
            }
            
        }

        /// <summary>
        /// Изменить константное лицо организации
        /// </summary>
        static void ChangeContactPerson()
        {
            if (!CheckRequiredTables("Клиенты"))
            {
                return;
            }

            if (dsMain.Tables["Клиенты"].Columns.IndexOf("Контактное лицо (ФИО)") == -1)
            {
                Console.WriteLine("В таблице \"Клиенты\" нет стоблца \"Контактное лицо (ФИО)\"");
                return;
            }

            PrintInfoFromTable("Клиенты");
            Console.WriteLine("Введите наименование организации для изменения контактного лица");
            string clientName = Console.ReadLine();
            dsMain.Tables["Клиенты"].PrimaryKey = new DataColumn[]{dsMain.Tables["Клиенты"].Columns["Наименование организации"] };
            DataRow findRow = dsMain.Tables["Клиенты"].Rows.Find(clientName);
            if (clientName != String.Empty && findRow != null)
            {
                Console.WriteLine("Введите новое контактное лицо");
                string newContactPerson = Console.ReadLine();
                if (newContactPerson != String.Empty)
                {
                    if (ChangeContactPersonInfile(newContactPerson, dsMain.Tables["Клиенты"].Rows.IndexOf(findRow)) == 1)
                    {
                        dsMain.Tables["Клиенты"].Rows[dsMain.Tables["Клиенты"].Rows.IndexOf(findRow)]["Контактное лицо (ФИО)"] = newContactPerson;
                        Console.WriteLine("Данные клиента успешно изменены");
                    }    
                }
            }
            else
            {
                Console.WriteLine("Клиент не найден");
            }
        }

        /// <summary>
        /// Изменение контактного лица в файле
        /// </summary>
        /// <param name="newContactPerson">Новое контактное лицо</param>
        /// <param name="rowNumber">Номер строки</param>
        /// <returns>Результат работы</returns>
        static int ChangeContactPersonInfile(string newContactPerson, int rowNumber)
        {
            try
            {
                using (XLWorkbook wBook = new XLWorkbook(fileExcel.FileStream))
                {
                    var workSheet = wBook.Worksheet("Клиенты");
                    workSheet.Cell(rowNumber+2, dsMain.Tables["Клиенты"].Columns.IndexOf("Контактное лицо (ФИО)")+1).Value = newContactPerson;//+2 т.к. первая строка заголовок
                    wBook.Save();
                    return 1;
                }
            }
            catch ( Exception ex)
            {
                Console.WriteLine("В результате изменения контактного лица клиента произошла ошибка"+ Environment.NewLine + ex.Message );
                return -1;
            }
        }

        /// <summary>
        /// Определение золотого клиента
        /// </summary>
        static void FindGoldenClient()
        {
            if (!CheckRequiredTables())
            {
                return;
            }

            Console.WriteLine("Введите поиск клиента за год или месяц\n 1 - текущий год\n 2 - определенный месяц");
            long chosenVariant = -1;
            if (!Int64.TryParse(Console.ReadLine(), out chosenVariant))
            {
                Console.WriteLine("Число не распознано");
                return;
            }

            long chosenMonth = -1;
            if (chosenVariant == 2)
            {
                Console.WriteLine("Введите номер месяца для вывода информации");
                if (!Int64.TryParse(Console.ReadLine(), out chosenMonth) && chosenMonth > 0 && chosenMonth < 13)
                {
                    Console.WriteLine("Некорректный номер месяца");
                    return;
                }
            }
            else if(chosenVariant != 1)
            {
                Console.WriteLine("Выбрана некорректная комманда");
                return;
            }

            var res = from req in dsMain.Tables["Заявки"].AsEnumerable()
                      join client in dsMain.Tables["Клиенты"].AsEnumerable() on req.Field<int>("Код клиента") equals client.Field<int>("Код клиента")
                      select new
                      { 
                          name = client.Field<string>("Наименование организации"),
                          countOfGoods = req.Field<int>("Требуемое количество"),
                          date = req.Field<DateTime>("Дата размещения")
                      };
            string goldenName = "";
            int goldenSumm = -1;
            if (chosenVariant == 1)
            {
                var res1 = (from table in res.AsEnumerable()
                            where table.date.Year == DateTime.Now.Year
                            group table by table.name into g
                            select new
                            {
                                clientId = g.Key,
                                sumOfGoods = g.Sum(x => x.countOfGoods)
                            }).OrderByDescending(x => x.sumOfGoods).Take(1);
                goldenName = res1.First().clientId;
                goldenSumm = res1.First().sumOfGoods;
            }
            else
            {
                var res1 = (from table in res.AsEnumerable()
                            where table.date.Year == DateTime.Now.Year && table.date.Month == chosenMonth
                            group table by new { clien = table.name, date = table.date } into g
                            select new
                            {
                                clientIdAndDate = g.Key,
                                sumOfGoods = g.Sum(x => x.countOfGoods),
                            }).OrderByDescending(x => x.sumOfGoods).Take(1);
                goldenName = res1.First().clientIdAndDate.clien;
                goldenSumm = res1.First().sumOfGoods;
            }


            if (!String.IsNullOrEmpty(goldenName))
            {

                Console.WriteLine($"Золотой клиент {goldenName} всего заказано на {(chosenVariant != 1 ? chosenMonth.ToString() + ".": "" )}" +
                    $"{DateTime.Now.Year} : {goldenSumm}"); 
            }
            else
            {
                Console.WriteLine("На выбранную дату нет заказов");
            }

        }
       
        
        #region Вспомогательные методы
                          /// <summary>
                          /// Загрузка инофрмации изи файла в DataSet
                          /// </summary>
                          /// <param name="path">Ппуть до файла (необязательный)</param>
        static int LoadInfoFromFile(string path = "excel.xlsx")
        {
            dsMain = new DataSet();
            string pathToFile = path;

            while (!File.Exists(pathToFile))
            {
                Console.WriteLine("Введите корректный путь до файла excel для открытия");
                pathToFile = Console.ReadLine();
            }
            try
            {
                fileExcel = new FileExcel(pathToFile, FileMode.OpenOrCreate);

                using (XLWorkbook wBook = new XLWorkbook(fileExcel.FileStream))
                {
                    var workSheets = wBook.Worksheets;
                    DataTable dt;
                    DataRow dr;
                    int countColumns;
                    int indexRow = 1;
                    foreach (var sheet in workSheets)
                    {
                        dt = new DataTable(sheet.Name);
                        for (int i = 1; i <= sheet.ColumnsUsed().Count(); i++)
                        {
                            dt.Columns.Add(sheet.Cell(1, i).GetText());
                            if (sheet.RowCount() > 1)
                            {
                                if (sheet.Cell(2, i).DataType == XLDataType.DateTime)
                                {
                                    dt.Columns[i - 1].DataType = typeof(DateTime);
                                }
                                else if (sheet.Cell(2, i).DataType == XLDataType.Number)
                                    dt.Columns[i - 1].DataType = typeof(int);
                            }

                        }
                        foreach (var row in sheet.RowsUsed())
                        {
                            if (row.RowNumber() != 1)
                            {
                                dr = dt.NewRow();
                                countColumns = sheet.ColumnsUsed().Count();
                                indexRow = 1;
                                foreach (var cell in row.CellsUsed())
                                {
                                    if (cell.DataType == XLDataType.DateTime)
                                        dr[indexRow - 1] = Convert.ToDateTime(cell.Value.ToString());
                                    else if (cell.DataType == XLDataType.Number)
                                        dr[indexRow - 1] = Convert.ToInt32(cell.Value.ToString());
                                    else
                                        dr[indexRow - 1] = cell.Value.ToString();
                                    indexRow++;
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                        dsMain.Tables.Add(dt);
                    }
                }

                CheckRequiredTables();
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка в процессе загрузки файла" + Environment.NewLine + ex.Message);
                return -1;
            }


        }

        /// <summary>
        /// Вывод информации из таблицы
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <returns>Код результата выполнения</returns>
        static int PrintInfoFromTable(string tableName)
        {
            if (dsMain.Tables.IndexOf(tableName) == -1)
            {
                Console.WriteLine($"В загруженном файле нет листа с именем \"{tableName}\"");
                return -1;
            }

            StringBuilder sb = new StringBuilder();
            foreach (DataColumn dc in dsMain.Tables[tableName].Columns)
            {
                sb.Append(String.Format("{0,-20}", dc.ColumnName));
            }
            sb.Append(Environment.NewLine);
            foreach (DataRow dr in dsMain.Tables[tableName].Rows)
            {
                for (int i = 0; i < dsMain.Tables[tableName].Columns.Count; i++)
                {
                    sb.Append(String.Format("{0,-20}", dr[i]));
                }
                sb.Append(Environment.NewLine);
            }
            Console.Write(sb.ToString());
            return 1;
        }

        /// <summary>
        /// Проверка наличия таблицы для работы программы
        /// </summary>
        /// <returns>Успешная првоерка</returns>
        static bool CheckRequiredTables()
        {
            bool flagSuccess = true;
            foreach (string tableName in requiredTables)
            {
                if (dsMain.Tables.IndexOf(tableName) == -1)
                {
                    Console.WriteLine($"В файле нет таблицы \"{tableName}\"");
                    flagSuccess = false;
                }
            }
            return flagSuccess;
        }

        /// <summary>
        /// Проверка наличия таблицы для работы программы
        /// </summary>
        /// <param name="tableName">Имя таблицы</param>
        /// <returns>Успешная првоерка</returns>
        static bool CheckRequiredTables(string tableName)
        {
            bool flagSuccess = true;
            if (dsMain.Tables.IndexOf(tableName) == -1)
            {
                Console.WriteLine($"В файле нет таблицы \"{tableName}\"");
                flagSuccess = false;
            }
            return flagSuccess;
        }
        #endregion
    }


    public abstract class FileBase
    { 
        public abstract string Path { get; set; }
        public abstract Stream FileStream { get; set; }
        public abstract void CreateStrema(FileMode fileMode);

    }

    public class FileExcel : FileBase
    {
        public override string Path { get; set; }
        public override Stream FileStream { get; set ; }

        public FileExcel(string path, FileMode fileMode = FileMode.Append)
        {
            Path = path;
            CreateStrema(fileMode);
        }

        public override void CreateStrema(FileMode fileMode = FileMode.Append)
        {
            FileStream = Path != "" ? new FileStream(Path, fileMode) : new FileStream("newFile.txt", fileMode);
        }
    }

}
