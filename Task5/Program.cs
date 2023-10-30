using System;
using System.Collections.Generic;
using System.Linq;

namespace Task5
{
    class Program
    {
        static void Main(string[] args)
        {
            var vacationDictionary = new Dictionary<string, List<DateTime>>()
            {
                ["Иванов Иван Иванович"] = new List<DateTime>(),
                ["Петров Петр Петрович"] = new List<DateTime>(),
                ["Юлина Юлия Юлиановна"] = new List<DateTime>(),
                ["Сидоров Сидор Сидорович"] = new List<DateTime>(),
                ["Павлов Павел Павлович"] = new List<DateTime>(),
                ["Георгиев Георг Георгиевич"] = new List<DateTime>()
            };
            var unAviableWorkingDaysOfWeekWithoutWeekends = new List<int>() { 6, 7 };
            // Список отпусков сотрудников
            List<DateTime> vacations = new List<DateTime>();
            List<DateTime> dateList = new List<DateTime>();
            List<DateTime> setDateList = new List<DateTime>();
            Random gen = new Random();
            Random step = new Random();

            DateTime start = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime end = new DateTime(start.Year, 12, 31);
            int range = (end - start).Days;
            int vacationCount = 28;
            DateTime startDate;
            DateTime endDate;
            int[] vacationSteps = { 7, 14 };
            int vacIndex = -1;
            int difference = -1;
            foreach (var vacatiomPerson in vacationDictionary)
            {
                vacationCount = 28;
                while (vacationCount > 0)
                {
                    startDate = start.AddDays(gen.Next(range));

                    if (!unAviableWorkingDaysOfWeekWithoutWeekends.Contains((int)startDate.DayOfWeek))
                    {
                        vacIndex = gen.Next(vacationSteps.Length);
                        endDate = new DateTime(DateTime.Now.Year, 12, 31);
                        difference = vacationSteps[vacIndex] == 7 ? 7 : 14;
                        endDate = startDate.AddDays(difference);

                        // Проверка условий по отпуску
                        if (!vacations.Any(x => startDate == x))
                        {
                            if (!vacations.Any(x => endDate == x))
                            {
                                for (DateTime dt = startDate; dt < endDate; dt = dt.AddDays(1))
                                {
                                    vacations.Add(dt);
                                    vacatiomPerson.Value.Add(dt);
                                }
                                vacationCount -= difference;
                            }
                        }
                    }
                }
            }
            foreach (var vacationList in vacationDictionary)
            {
                vacationList.Value.Sort();
                Console.WriteLine("Дни отпуска " + vacationList.Key + " : ");
                foreach (var date in vacationList.Value)
                {
                    Console.WriteLine(date.ToShortDateString());
                }
            }
            Console.ReadKey();
        }
    }
}