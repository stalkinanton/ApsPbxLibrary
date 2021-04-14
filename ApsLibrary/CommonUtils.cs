using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Preactor;
using Preactor.Interop.PreactorObject;
using static ApsPbxLibrary.LogFile;

namespace ApsPbxLibrary
{
    public static class CU
    {
        /// <summary>
        /// Значение даты "Не указано" в Preactor
        /// </summary>
        public static DateTime UNSPECIFIED_DATE = DateTime.FromOADate(-1);

        public static IPreactor Preactor;
        public static IEventScriptsCore Core;

        static TimeSpan? schedulingAccuracy = null;
        public static TimeSpan SCHEDULING_ACCURACY
        {
            get
            {
                if (!schedulingAccuracy.HasValue)
                    schedulingAccuracy = TimeSpan.FromDays(Preactor.PlanningBoard.SchedulingAccuracy);
                return schedulingAccuracy.Value;
            }
        }

        /// <summary>
        /// Перевод даты в формат, нужный в поисковых выражениях Preactor
        /// </summary>
        /// <param name="date">дата</param>
        /// <returns>Дата в строковом формате</returns>
        public static string DateToExprStr(DateTime date)
        {
            return date.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        public enum SmcRule
        {
            Default
        }

        /// <summary>
        /// Связка материалов по Default
        /// </summary>
        /// <param name="rule">используемое правило связки</param>
        public static void RunSmc(SmcRule rule)
        {
            Dictionary<SmcRule, string> rulePespNames = new Dictionary<SmcRule, string>()
            { { SmcRule.Default, "SMC Default" } };

            if (!rulePespNames.ContainsKey(rule))
                throw new NotSupportedException(rule.ToString());

            Preactor.RunEventScript(rulePespNames[rule]);
        }

        /// <summary>
        /// Первый день месяца
        /// </summary>
        /// <param name="date">произвольная дата в рамках месяца</param>
        /// <returns>Первый день месяца</returns>
        public static DateTime FirstDayOfMonth(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

#region Перегрузки метода FindMatchingRecords

        /// <summary>
        /// Поиск всех записей в заданной таблице, которые соответствуют условию
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="searchExpr">условие поиска</param>
        /// <returns>Список номеров записей</returns>
        public static List<int> FindMatchingRecords(string tableName, string searchExpr)
        {
            if (Preactor == null)
                Log("Поле Preactor в CommonUtils не инициализировано", LogSeverity.Error);

            List<int> res = new List<int>();
            int recNo = Preactor.FindMatchingRecord(tableName, 0, searchExpr);
            while (recNo > 0)
            {
                res.Add(recNo);
                recNo = Preactor.FindMatchingRecord(tableName, recNo, searchExpr);
            }
            return res;
        }

        /// <summary>
        /// Поиск всех записей в заданной таблице по значению поля
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="fieldName">имя поля</param>
        /// <param name="fieldValue">требуемое значение</param>
        /// <returns>Список номеров записей</returns>
        public static List<int> FindMatchingRecords(string tableName, string fieldName, string fieldValue)
        {
            if (Preactor == null)
                Log("Поле Preactor в CommonUtils не инициализировано", LogSeverity.Error);

            List<int> res = new List<int>();
            int recNo = Preactor.FindMatchingRecord(tableName, fieldName, 0, fieldValue);
            while (recNo > 0)
            {
                res.Add(recNo);
                recNo = Preactor.FindMatchingRecord(tableName, fieldName, recNo, fieldValue);
            }
            return res;
        }

        /// <summary>
        /// Поиск всех записей в заданной таблице по значению поля
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="fieldName">имя поля</param>
        /// <param name="fieldValue">требуемое значение</param>
        /// <returns>Список номеров записей</returns>
        public static List<int> FindMatchingRecords(string tableName, string fieldName, bool fieldValue)
        {
            if (Preactor == null)
                Log("Поле Preactor в CommonUtils не инициализировано", LogSeverity.Error);

            List<int> res = new List<int>();
            int recNo = Preactor.FindMatchingRecord(tableName, fieldName, 0, fieldValue);
            while (recNo > 0)
            {
                res.Add(recNo);
                recNo = Preactor.FindMatchingRecord(tableName, fieldName, recNo, fieldValue);
            }
            return res;
        }

        /// <summary>
        /// Поиск всех записей в заданной таблице по значению поля
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="fieldName">имя поля</param>
        /// <param name="fieldValue">требуемое значение</param>
        /// <returns>Список номеров записей</returns>
        public static List<int> FindMatchingRecords(string tableName, string fieldName, DateTime fieldValue)
        {
            if (Preactor == null)
                Log("Поле Preactor в CommonUtils не инициализировано", LogSeverity.Error);

            List<int> res = new List<int>();
            int recNo = Preactor.FindMatchingRecord(tableName, fieldName, 0, fieldValue);
            while (recNo > 0)
            {
                res.Add(recNo);
                recNo = Preactor.FindMatchingRecord(tableName, fieldName, recNo, fieldValue);
            }
            return res;
        }

        /// <summary>
        /// Поиск всех записей в заданной таблице по значению поля
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="fieldName">имя поля</param>
        /// <param name="fieldValue">требуемое значение</param>
        /// <returns>Список номеров записей</returns>
        public static List<int> FindMatchingRecords(string tableName, string fieldName, int fieldValue)
        {
            if (Preactor == null)
                Log("Поле Preactor в CommonUtils не инициализировано", LogSeverity.Error);

            List<int> res = new List<int>();
            int recNo = Preactor.FindMatchingRecord(tableName, fieldName, 0, fieldValue);
            while (recNo > 0)
            {
                res.Add(recNo);
                recNo = Preactor.FindMatchingRecord(tableName, fieldName, recNo, fieldValue);
            }
            return res;
        }

        /// <summary>
        /// Поиск всех записей в заданной таблице по значению поля
        /// </summary>
        /// <param name="tableName">имя таблицы</param>
        /// <param name="fieldName">имя поля</param>
        /// <param name="fieldValue">требуемое значение</param>
        /// <returns>Список номеров записей</returns>
        public static List<int> FindMatchingRecords(string tableName, string fieldName, double fieldValue)
        {
            if (Preactor == null)
                Log("Поле Preactor в CommonUtils не инициализировано", LogSeverity.Error);

            List<int> res = new List<int>();
            int recNo = Preactor.FindMatchingRecord(tableName, fieldName, 0, fieldValue);
            while (recNo > 0)
            {
                res.Add(recNo);
                recNo = Preactor.FindMatchingRecord(tableName, fieldName, recNo, fieldValue);
            }
            return res;
        }

#endregion

#region Перегрузки метода WriteScriptVariable

        /// <summary>
        /// Создание PESP-переменной с заданным значением
        /// </summary>
        /// <param name="name">имя переменной</param>
        /// <param name="value">значение</param>
        public static void WriteScriptVariable(string name, string value)
        {
            if (Core == null)
                Log("Поле Core в CommonUtils не инициализировано", LogSeverity.Error);
            Core.WriteScriptVariable(name, value);
        }

        /// <summary>
        /// Создание PESP-переменной с заданным значением
        /// </summary>
        /// <param name="name">имя переменной</param>
        /// <param name="value">значение</param>
        public static void WriteScriptVariable(string name, bool value)
        {
            if (Core == null)
                Log("Поле Core в CommonUtils не инициализировано", LogSeverity.Error);
            Core.WriteScriptVariable(name, value);
        }

        /// <summary>
        /// Создание PESP-переменной с заданным значением
        /// </summary>
        /// <param name="name">имя переменной</param>
        /// <param name="value">значение</param>
        public static void WriteScriptVariable(string name, DateTime value)
        {
            if (Core == null)
                Log("Поле Core в CommonUtils не инициализировано", LogSeverity.Error);
            Core.WriteScriptVariable(name, value);
        }

        /// <summary>
        /// Создание PESP-переменной с заданным значением
        /// </summary>
        /// <param name="name">имя переменной</param>
        /// <param name="value">значение</param>
        public static void WriteScriptVariable(string name, double value)
        {
            if (Core == null)
                Log("Поле Core в CommonUtils не инициализировано", LogSeverity.Error);
            Core.WriteScriptVariable(name, value);
        }

        /// <summary>
        /// Создание PESP-переменной с заданным значением
        /// </summary>
        /// <param name="name">имя переменной</param>
        /// <param name="value">значение</param>
        public static void WriteScriptVariable(string name, int value)
        {
            if (Core == null)
                Log("Поле Core в CommonUtils не инициализировано", LogSeverity.Error);
            Core.WriteScriptVariable(name, value);
        }

#endregion
    }

    /// <summary>
    /// Класс исключения для плановой остановки выполнения библиотеки в любой точке
    /// </summary>
    public class AbortException: Exception { }

    /// <summary>
    /// Прогресс операции
    /// </summary>
    public enum OperationProgress
    {
        NotStarted = 2,
        Setup = 3,
        Running = 4,
        Complete = 5
    }

    /// <summary>
    /// Тип времени процесса
    /// </summary>
    public enum ProcessTimeType
    {
        TimePerItem = 0,
        TimePerBatch = 1,
        RatePerHour = 2,
        ResourceTimePerItem = 3,
        ResourceTimePerBatch = 5,
        ResourceRatePerHour = 4,
    }
}
