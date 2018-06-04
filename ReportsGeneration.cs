    using System;
    using System.Collections.Generic;
    using OfficeOpenXml;
    using System.IO;
    using System.Data;
    using System.Linq;

namespace ReportGenerations
{/// <summary>
 /// Клас генерации отчетов
 /// </summary>
    public static class ReportsGeneration
    {
        /// <summary>
        /// DataSetTable с данными и переменными values
        /// </summary>
        static DataSetTable dST;
        /// <summary>
        /// Массив байт с результатами
        /// </summary>
        static byte[] outresukt;
        /// <summary>
        /// Метод генерации отчетов
        /// </summary>
        /// <param name="templateFile">Путь к excel шаблону отчета</param>
        /// <param name="_dST">DataSetTable с данными  и переменными values</param>
        /// <param name="destinationFile">Путь c именем результата</param>
        /// <param name="toByte">Если true, то возвращает</param>
        public static void ReportGeneration(string templateFile, DataSetTable _dST, string destinationFile, bool toByte = false)
        {
            dST = _dST;
            using (ExcelPackage p = new ExcelPackage(new FileInfo(templateFile), true))
            {
                using (ExcelWorksheet ws = p.Workbook.Worksheets[1])
                {
                    var startCountRows = 0;
                    for (int indexRow = 1; indexRow <= ws.Dimension.Rows; indexRow++)
                    {
                        for (int indexColum = 1; indexColum <= ws.Dimension.Columns; indexColum++)
                        {
                            var cellValue = ws.Cells[indexRow, indexColum].Value;//типо оптимизация
                                                                                 //  ref object CellValue =ref ws.Cells[indexRow, indexColum].Value;//создаем ссылку на значение ячейки, но только в новом шарпе

                            //именно проверяем на null что б не падал, а на пусто ту не проверяем, т.к. следующие условия его обрабатывают
                            if (cellValue == null)
                            {
                                continue;
                            }
                            else if (cellValue.ToString().Contains("<sub>"))//дефолтный под итог, потом можно расширить на не дублирование и подведение итогов не в одной колонке
                            {
                                ws.Cells[indexRow, indexColum].Value = cellValue.ToString().Replace("<sub>", "");//очищаем от тех. инфы
                                startCountRows = indexRow;
                                string tableName = cellValue.ToString().Replace("${", "").Replace("}", "");
                                tableName = tableName.Substring(0, tableName.IndexOf("."));//вычленяем имя
                                var items = dST.Tables[tableName].Rows;//берем строки из таблицу с нужным названием

                                List<subStrucktur> subStrucktures = new List<subStrucktur>();//хранилище структуры подитогов
                                int nextRow = 1;
                                int countSub = 0;
                                while (true)
                                {
                                    var valueCell = ws.Cells[indexRow + nextRow, indexColum].Value?.ToString();//получаем следующую ячейку
                                    if (valueCell == null)
                                        if (!(subStrucktures.Count > 0 && subStrucktures.Count == countSub)) nextRow++;//если еще не просмотрели всю структуру, то продолжаем
                                        else break;
                                    else if (valueCell.Contains("<subitog"))//Проверяем есть ли ключевое слово
                                    {
                                        if (countSub == 0)//если количество итогов не задано, парсим
                                            countSub = int.Parse(valueCell.Substring(valueCell.IndexOf(":") - 1, 1));//вынимаем кол-во итогов
                                                                                                                     //добываем ссылку на данные
                                        var tempValue = valueCell.Substring(valueCell.IndexOf(":") + 1);
                                        tempValue = tempValue.Substring(0, tempValue.IndexOf(">")).Replace("${", "");
                                        tempValue = tempValue.Substring(tempValue.IndexOf(".") + 1).Replace("}", "");
                                        if (subStrucktures.Count > 0 && nextRow - 1 != subStrucktures.Last().NumberRow)
                                            subStrucktures.Last().nextRowEmpty = true;//флаг что после итога нужна пустая строка
                                        subStrucktures.Add(new subStrucktur(tempValue, nextRow, tempValue == "full" ? "full" : items[0][tempValue].ToString(), startCountRows) { Row = items[0] });

                                        ws.Cells[indexRow + nextRow, indexColum].Value = valueCell.Remove(0, valueCell.IndexOf(">") + 1);//вырезаем тех.инфу
                                        nextRow++;
                                    }
                                    else nextRow++;
                                }
                                foreach (DataRow item in items)//построчная вставка данных, построчная, что бы было проше анализировать формулы и тп.
                                {
                                    foreach (var pZnach in subStrucktures)//проверяем сформировались ли подитоги
                                    {
                                        if (pZnach.Key != "full")
                                        {
                                            ProverkaItogov(ws, ref indexRow, tableName, pZnach, item);
                                        }
                                    }
                                    InsertData(ws, ref indexRow, tableName, item);//вставляем строку данных с обработкой формул
                                }
                                //что бы подвести промежуточные итоги последней строке и итоговые
                                foreach (var pZnach in subStrucktures)
                                {
                                    ProverkaItogov(ws, ref indexRow, tableName, pZnach, dST.Tables[tableName].NewRow());
                                }
                                //    ProverkaItogov(ws, ref indexRow, tableName, subStrucktures.Last(),  dST.Tables[tableName].NewRow());//старые итоговые
                                ws.DeleteRow(indexRow);//удаляем строку шаблона

                                for (int i = 0; i < subStrucktures.Count; i++)//строки итогов и саму строку удалить
                                {
                                    if (subStrucktures[i].nextRowEmpty)
                                    {
                                        ws.DeleteRow(indexRow);//удаляем пустую строку шаблона подитогов
                                    }
                                    ws.DeleteRow(indexRow);//удаляем строку шаблона подитогов
                                }
                                for (int j = 1; j <= ws.Dimension.Columns; j++)//проходим по всем столбцам проверяем именнованные ряды
                                {
                                    ReDiapozonNameRange(ws, 0, indexRow, j, startCountRows);
                                }
                                break;//так как вставлялка данных проверяет все колонки, то индекс столбца надо сбросить
                            }
                            else if (cellValue.ToString().Contains("${"))
                            {
                                if (cellValue.ToString().Contains("${graf"))//todo реализовать метки для нескольких графиков думаю через заголовок графика
                                {
                                    var name = cellValue.ToString().Substring(cellValue.ToString().IndexOf(".") + 1);
                                    name = name.Substring(0, name.IndexOf("}"));//имя можно посмотреть там же где задается именнованный ряд, изменить имя можно на вкладке графика
                                    ws.Cells[indexRow, indexColum].Value = null;
                                    var tempChar = ws.Drawings[name] as ExcelChart;
                                        tempChar.SetPosition(indexRow - 1, 0, indexColum - 1, 0);//двигаем график на нужное место

                                    while (tempChar.Title.Text.Contains("${"))//обработка нескольких переменных в одной ячейке
                                    {
                                        var templateValue = tempChar.Title.Text.Substring(tempChar.Title.Text.IndexOf("${"));
                                        templateValue = templateValue.Substring(0, templateValue.IndexOf("}") + 1);
                                        tempChar.Title.Text = cellValue.ToString().Replace(templateValue, (string)dST.values[templateValue.Replace("${", "").Replace("}", "")]);
                                    }
                                  }
                                else if (!cellValue.ToString().Substring(cellValue.ToString().IndexOf("${")).Contains("."))//todo в др метса
                                {
                                    var tempValue = cellValue.ToString();
                                    while (tempValue.Contains("${"))//обработка нескольких переменных в одной ячейке
                                    {
                                        var templateValue = tempValue.Substring(tempValue.IndexOf("${"));
                                        templateValue = templateValue.Substring(0, templateValue.IndexOf("}") + 1);
                                        tempValue = cellValue.ToString().Replace(templateValue, (string)dST.values[templateValue.Replace("${", "").Replace("}", "")]);
                                    }
                                    ws.Cells[indexRow, indexColum].Value = tempValue;
                                }
                                else
                                {
                                    startCountRows = indexRow;//запоминаем начало вставки данных
                                    string tableName = cellValue.ToString().Replace("${", "").Replace("}", "");
                                    tableName = tableName.Substring(0, tableName.IndexOf("."));//вычленяем имя
                                    var items = dST.Tables[tableName].Rows;   //берем строки из таблицу с нужным названием                             
                                    foreach (DataRow item in items)//построчная вставка данных, построчная, что бы было проше анализировать формулы и тп.
                                    {
                                        InsertData(ws, ref indexRow, tableName, item);//вставляем строку данных с обработкой формул
                                    }
                                    for (int j = 1; j <= ws.Dimension.Columns; j++)//проходим по всем столбцам проверяем именнованные ряды
                                    {
                                        ReDiapozonNameRange(ws, 0, indexRow, j, startCountRows);
                                    }
                                    /*  foreach (var nameRange in p.Workbook.Names)
                                      {
                                          nameRange.Address = ExcelCellBase.GetAddress(startCountRows, nameRange.Start.Column, startCountRows + countItem, nameRange.End.Column);
                                      }*/
                                    ws.DeleteRow(indexRow);//удаляем строку шаблона
                                    indexRow--;//удалили строку, уменьшаем индекс текущей строки
                                    break;//так как вставлялка данных проверяет все колонки, то индекс надо сбросить
                                }
                            }
                            else if (cellValue.ToString().IndexOf("$[") != -1)//обработка формул
                                ws.Cells[indexRow, indexColum].Formula = cellValue.ToString().Replace("@", indexRow.ToString()).Replace("$[", "").Replace("]", "").Replace(";", ",");
                            else continue;
                        }
                    }
                    ws.Calculate();//просчитываем формулы
                    if (!toByte)
                        try
                        {
                            p.SaveAs(new FileInfo(destinationFile));//сохраняем файл
                        }
                        catch (Exception)
                        {   //если файл открыть или нельзя редактировать, сохраняем файл с указанием времени
                            p.SaveAs(new FileInfo(destinationFile.Insert(destinationFile.LastIndexOf("."), DateTime.Now.ToString("yyyyMMddHHmmss"))));
                        }
                    else
                        outresukt = p.GetAsByteArray();//кому надо можно получить байты, пока жив класс, правда я пока не тестил
                }
            }
        }
        /// <summary>
        ////Вставляет строку данных в шаблон
        /// </summary>
        /// <param name="ws">ExcelWorksheet</param>
        /// <param name="indexRow">номер строки начала и потом окончания вставки данных</param>
        /// <param name="tableName">название таблицы</param>
        /// <param name="item">строка данных</param>
        private static void InsertData(ExcelWorksheet ws, ref int indexRow, string tableName, DataRow item)
        {
            int j;
            ws.InsertRow(indexRow, 1, indexRow + 1);//вставляем строку
            ws.Row(indexRow).Height = ws.Row(indexRow + 1).Height;
            for (j = 1; j <= ws.Dimension.Columns; j++)//где то здесь что то пихнуть ,что бы информация (по которой подводятся итоги) не дублировалась по строкам
            {
                var cellValue = ws.Cells[indexRow + 1, j].Value?.ToString();//берем значение ячейки с учетом, что может быть нулл
                if (ws.Cells[indexRow + 1, j].Value == null)
                    continue;
                else if (cellValue.Contains("${"))
                    if (cellValue.Substring(cellValue.IndexOf("${")).Contains("."))//смотрим это ссылка на данные?
                    {
                        var keyValue = cellValue.Substring(cellValue.IndexOf("${"));
                        keyValue = keyValue.Substring(0, keyValue.IndexOf("}") + 1);
                        /* double res;
                         var value = cellValue.Replace(keyValue, item[keyValue.Replace($"${{{tableName}.", "").Replace("}", "")].ToString());
                         ws.Cells[indexRow, j].Value = double.TryParse(value, out res) ? (object)res : (object)value;
                         */
                        //tryparse для того что бы числа вставлялись как числа, даты как даты, можно дополнить...
                        ws.Cells[indexRow, j].Value = (cellValue.Replace(keyValue, item[keyValue.Replace($"${{{tableName}.", "").Replace("}", "")].ToString())).TryParse<DateTime, double>();
                    }
                    else
                    {
                        var tempValue = cellValue.ToString();
                        while (tempValue.Contains("${"))//обрабатываем переменные
                        {
                            var templateValue = tempValue.Substring(tempValue.IndexOf("${"));
                            templateValue = templateValue.Substring(0, templateValue.IndexOf("}") + 1);
                            tempValue = tempValue.Replace(templateValue, (string)dST.values[templateValue.Replace("${", "").Replace("}", "")]);
                        }
                        ws.Cells[indexRow, j].Value = tempValue;
                    }
                else if (cellValue.Contains("$["))//обработка формул
                    ws.Cells[indexRow, j].Formula = cellValue.Replace("@", indexRow.ToString()).Replace("$[", "").Replace("]", "").Replace(";", ",");
                else//просто значение
                    ws.Cells[indexRow, j].Value = ws.Cells[indexRow + 1, j].Value;
            }
            indexRow++;
        }

        /// <summary>
        /// Проводим проверку не пора ли вставить под итог
        /// </summary>
        /// <param name="ws">ExcelWorksheet</param>
        /// <param name="indexRow">индекс строки в котором начали проверять и закончили</param>
        /// <param name="tableName">название таблицы которую используем из дата сет</param>
        /// <param name="subStrucktures">элемент подитогов для проверки</param>
        /// <param name="j">столбец в котором начали проверять и закончили</param>
        /// <param name="item">строка данных</param>
        private static void ProverkaItogov(ExcelWorksheet ws, ref int indexRow, string tableName, subStrucktur pZnach, DataRow item)
        {
            var znachItem = pZnach.Key != "full" ? item[pZnach.Key].ToString() : "";
            if (pZnach.PrevZnach != znachItem)//о значение которое мы отслеживали изменилось, подводим под итог
            {
                ws.InsertRow(indexRow, 1, indexRow + 1 + pZnach.NumberRow);//вставка строки (куда, количество, с какой строки копировать формат)
                ws.Row(indexRow).Height = ws.Row(indexRow + 1 + pZnach.NumberRow).Height;//высота строки не является формавтом, руками изменяем
                for (int j = 1; j <= ws.Dimension.Columns; j++)//где то здесь что то пихнуть ,что бы информация не дублировалась по строкам
                {
                    var cellValue = ws.Cells[indexRow + 1 + pZnach.NumberRow, j].Value?.ToString();
                    if (cellValue == null)
                        continue;
                    else if (cellValue.Contains("${"))
                    {//вставка значения из таблицы
                        if (cellValue.Substring(cellValue.IndexOf("${")).Contains("."))
                        {
                            var keyValue = cellValue.Substring(cellValue.IndexOf("${"));
                            keyValue = keyValue.Substring(0, keyValue.IndexOf("}") + 1);
                            //tryparse для того что бы числа вставлялись как числа, даты как даты, можно дополнить...
                            ws.Cells[indexRow, j].Value = (cellValue.Replace(keyValue, pZnach.Row[keyValue.Replace($"${{{tableName}.", "").Replace("}", "")].ToString())).TryParse<DateTime, double>();
                        }// pZnach.Row[cellValue.Replace(tableName + ".", "").Replace("${", "").Replace("}", "")]; }
                        else
                        {//вставка переменных
                            var tempValue = cellValue.ToString();
                            while (tempValue.Contains("${"))
                            {
                                var templateValue = tempValue.Substring(tempValue.IndexOf("${"));
                                templateValue = templateValue.Substring(0, templateValue.IndexOf("}") + 1);
                                tempValue = tempValue.Replace(templateValue, (string)dST.values[templateValue.Replace("${", "").Replace("}", "")]);
                            }
                            ws.Cells[indexRow, j].Value = tempValue;
                        }

                    }
                    else if (cellValue.Contains("$["))//обработка формул
                    {
                        ws.Cells[indexRow, j].Formula = cellValue.Replace("@", indexRow.ToString()).Replace("$[", "").Replace("]", "").Replace(";", ",");
                    }
                    else if (cellValue.Contains("<"))//обработка формул под итогов, пока только сумма
                    {
                        int oper = cellValue == "<SUM>" ? 9 : 0;
                        ws.Cells[indexRow, j].Formula = $"SUBTOTAL({oper}, {ExcelCellBase.GetAddress(pZnach.startRowBlock, j, indexRow - 1, j)})";
                    }
                    else
                    {
                        ws.Cells[indexRow, j].Value = cellValue;
                    }
                    if (pZnach.Key == "full")//если это итоговый подитог пересчитать именнованные ячейки
                    {
                        ReDiapozonNameRange(ws, 1 + pZnach.NumberRow, indexRow, j);
                    }
                }

                indexRow++;
                if (pZnach.nextRowEmpty)//если после подитога нужна пустая строка, вставляем ее
                {
                    ws.InsertRow(indexRow, 1, indexRow + 1 + pZnach.NumberRow + 1);
                    ws.Row(indexRow).Height = ws.Row(indexRow + 1 + pZnach.NumberRow + 1).Height;
                    indexRow++;//мы вставили строку, надо это учитывать
                }
                //обновляем метки
                pZnach.PrevZnach = znachItem;
                pZnach.startRowBlock = indexRow;
                pZnach.Row = item;// у меня извращенный скрипт, и по этому я сохраняю первую строку нового блока
            }
        }

        /// <summary>
        /// Обновление адресов именнованных рядов
        /// </summary>
        /// <param name="ws">ExcelWorksheet</param>
        /// <param name="smesh">смешение от текущей строки где надо искать подитог</param>
        /// <param name="tindexRow">номер строки</param>
        /// <param name="tJ">номер столбца</param>
        /// <param name="startCountRows">начало вставки данных</param>
        private static void ReDiapozonNameRange(ExcelWorksheet ws, int smesh, int tindexRow, int tJ, int startCountRows = 0)
        {
            try
            {//находим соответствие и обновляем диапозон
             //  ws.Workbook.Names.First(x => x.Address == ws.Cells[tindexRow + smesh, tJ].FullAddress).Address = ws.Cells[tindexRow, tJ].FullAddress;
                ws.Workbook.Names.First(x => x.End.Address == ws.Cells[tindexRow + smesh, tJ].FullAddress || x.End.Address == ws.Cells[tindexRow + smesh, tJ].Address).Address = ExcelCellBase.GetAddress(startCountRows == 0 ? tindexRow : startCountRows, tJ, tindexRow, tJ);
            }
            catch (Exception ee)
            {
                // Console.WriteLine("При переопределении диапазона именованных ячеек произошла ошибка: " + ee.Message);
            }
        }
    }
    /// <summary>
    /// Класс для подитогов
    /// </summary>
    public class subStrucktur
    {
        /// <summary>
        /// Наименование столбца, по которому отслеживаем изменение
        /// </summary>
        public string Key { get; private set; }
        /// <summary>
        /// номер строки после даных(начинается с 1)
        /// </summary>
        public int NumberRow { get; private set; }
        /// <summary>
        /// Предыдущее значение столбца, по которому отслеживаем изменение
        /// </summary>
        public string PrevZnach { get; internal set; }
        /// <summary>
        /// флаг следующая строка после под итога пустая
        /// </summary>
        public bool nextRowEmpty { get; internal set; }
        /// <summary>
        /// номер строки с которого начался блок данных, пока не используется
        /// </summary>
        public int startRowBlock { get; internal set; }
        /// <summary>
        /// первая строка данных следующего блока
        /// </summary>
        public DataRow Row { get; internal set; }
        /// <summary>
        /// экземпляр подитогов
        /// </summary>
        /// <param name="_key">Наименование колонки по которой отслеживается изменения</param>
        /// <param name="_numberRow">номер строки от данных, на которой распологается подитог</param>
        /// <param name="_prevZnach">предыдущее значение ключа</param>
        /// <param name="_startRowBlock">номер строки с которой начался текущий блок для подитогов</param>
        /// <param name="_nextRowEmpty">следующая строка после подитогов пустая?</param>
        public subStrucktur(string _key, int _numberRow, string _prevZnach = "", int _startRowBlock = 0, bool _nextRowEmpty = false)
        {
            Key = _key;
            NumberRow = _numberRow;
            PrevZnach = _prevZnach;
            startRowBlock = _startRowBlock;
            nextRowEmpty = _nextRowEmpty;
        }
    }
    /// <summary>
    /// Бонус, класс помошник
    /// </summary>
    static class help
    {
        /// <summary>
        /// Пытается конвертировать в первый тип, если не получается во второй, иначе возвращает стринг
        /// </summary>
        /// <typeparam name="T">Тип во что конвертируем (дата)</typeparam>
        /// <typeparam name="U">Тип во что конвертируем (число)</typeparam>
        /// <param name="value">Что конвертируем</param>
        /// <returns></returns>
        public static object TryParse<T, U>(this string value)
        {
            // T res;
            try
            {
                var converter = System.ComponentModel.TypeDescriptor.GetConverter(typeof(T));
                if (converter != null)
                {
                    // Cast ConvertFromString(string text) : object to (T)
                    return (T)converter.ConvertFromString(value);
                }
                //  return value;
            }
            catch (Exception) { }
            try
            {
                var converter = System.ComponentModel.TypeDescriptor.GetConverter(typeof(U));
                if (converter != null)
                {
                    // Cast ConvertFromString(string text) : object to (T)
                    return (U)converter.ConvertFromString(value);
                }
                return value;
            }
            catch (Exception)
            {
                return value;
            }
        }
        /// <summary>
        /// Пытается конвертировать в заданный тип, иначе возвращает стринг
        /// </summary>
        /// <typeparam name="T">Тип во что конвертируем</typeparam>
        /// <param name="value">Что конвертируем</param>
        /// <returns></returns>
        public static object TryParse<T>(this string value)
        {
            // T res;
            try
            {
                var converter = System.ComponentModel.TypeDescriptor.GetConverter(typeof(T));
                if (converter != null)
                {
                    // Cast ConvertFromString(string text) : object to (T)
                    return (T)converter.ConvertFromString(value);
                }
                return value;
            }
            catch (Exception e)
            {
                return value;
            }
            // return T.TryParse(value, out res)?(object)res:(object)value;
        }
        /// <summary>
        /// преобразование числа.ToString() в палочки
        /// </summary>
        /// <param name="countShare">число в стринге</param>
        /// <returns>палочки</returns>
        public static string getShare(this string countShare)
        {
            string sh = "";
            for (int i = 0; i < int.Parse(countShare); i++) sh += "|";
            return sh;
        }
    }
}
