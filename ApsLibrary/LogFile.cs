using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace ApsPbxLibrary
{
    public static class LogFile
    {
        static TextWriter writer;
        static List<string> warnings;

        /// <summary>
        /// Инициализация журнала
        /// </summary>
        /// <param name="fileName">путь к фалу журнала</param>
        public static void Init(string fileName)
        {
            try
            {
                writer = new StreamWriter(fileName);
                warnings = new List<string>();
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                MessageBox.Show("Не найден путь к файлу журнала.\r\nЖурнал не будет вестись.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Закрытие файла журнала
        /// </summary>
        public static void Close()
        {
            const int MAX_ELEMS_IN_MESSAGE = 20;

            if (writer != null)
            {
                writer.Dispose();
                writer = null;
            }
            if (warnings.Any())
            {
                string str = "В ходе выполнения обнаружены предупреждения:";
                if (warnings.Count <= MAX_ELEMS_IN_MESSAGE)
                {
                    warnings.ForEach(w => str += "\n" + w);
                }
                else
                {
                    warnings.Take(MAX_ELEMS_IN_MESSAGE).ToList().ForEach(w => str += "\n" + w);
                    str += $"\n\nПолный список предупреждений записан в лог файл ({warnings.Count()} записей).";
                }
                MessageBox.Show(str, "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            warnings = null;
        }

        /// <summary>
        /// Внесение записи в журнал
        /// </summary>
        /// <param name="str">вносимая строка</param>
        /// <param name="severity">уровень серьезности сообщения</param>
        public static void Log(string str, LogSeverity severity = LogSeverity.Info)
        {
            if (severity == LogSeverity.Warning)
            {
                if ((warnings != null) && !warnings.Contains(str))
                    warnings.Add(str);
                str = "WARNING: " + str;
            }
            if (severity == LogSeverity.Error)
            {
                str = "ERROR: " + str;
            }
            if (writer != null)
            {
                writer.WriteLine(str);
                writer.Flush();
            }
            if (severity == LogSeverity.Error)
            {
                throw new ApplicationException(str.Substring(7, str.Length - 7));
            }
        }

        /// <summary>
        /// Уровень серьезности сообщения журнала
        /// </summary>
        public enum LogSeverity
        {
            /// <summary>Информационное сообщение, фиксируется только в журнал</summary>
            Info,
            /// <summary>Предупреждение, выдается пользователю с воскл. знаком в желтом треугольнике, но не останавливает работу</summary>
            Warning,
            /// <summary>Ошибка, выдается с красным крестиком и останавливает работу алгоритма</summary>
            Error
        }
    }
}
