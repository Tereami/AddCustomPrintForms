#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-NonCommercial-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в некоммерческих целях,
при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-NonCommercial-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion
using System;
using System.Linq;
using System.Text;
using System.IO;

namespace AddCustomPrintForms
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Запуск");

            string path = @"\\HKEY_LOCAL_MACHINE\SOFTWARE\RevitBatchPrint\WereCustomFormsAdded";
            string batchPrintKey = "RevitBatchPrint";
            string registryKey = "WereCustomFormsAdded";

            Console.WriteLine("Проверяем, не запускался ли скрипт ранее, ключ реестра " + path);

            Microsoft.Win32.RegistryKey localMachineKey =
                Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64);
                //Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey softwareSubkey = localMachineKey.OpenSubKey("SOFTWARE", true);
            Microsoft.Win32.RegistryKey batchPrintSubkey = null;

            if (softwareSubkey.GetSubKeyNames().Contains(batchPrintKey))
            {
                batchPrintSubkey = softwareSubkey.OpenSubKey(batchPrintKey, true);
            }
            else
            {
                batchPrintSubkey = softwareSubkey.CreateSubKey(batchPrintKey);
            }

            object checkKeyExists = batchPrintSubkey.GetValue(registryKey);
            if (checkKeyExists == null )
            {
                Console.WriteLine("Запущен впервые.");
            }
            else
            {
                if (checkKeyExists.ToString() == "true")
                {
                    Console.WriteLine("Запущен повторно. Форматы будут перезаписаны.");
                }
                else
                {
                    Console.WriteLine("Уже был запущен, но отработал некорректно. Будет запущен заново.");
                }
            }
            batchPrintSubkey.SetValue(registryKey, "error");

            string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            
            string collectionFilename = Path.Combine(Path.GetDirectoryName(assemblyPath), "formats.txt");

            bool checkFile = File.Exists(collectionFilename);
            if(checkFile)
            {
                Console.WriteLine("Чтение списка форматов");
            }
            else
            {
                Console.WriteLine("Не найден файл formats.txt");
                Console.ReadKey();
                return;
            }

            string[] lines = File.ReadAllLines(collectionFilename, Encoding.UTF8);

            string printerName = "";
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                Console.WriteLine("Проверяется принтер: " + printer);
                if (printer == "PDFCreator" || printer.Contains("PDFCreator"))
                {
                    Console.WriteLine("Найден принтер PDFCreator");
                    printerName = "PDFCreator";
                    break;
                }
            }
            if(printerName == "")
            {
                Console.WriteLine("Принтер PDFCreator не установлен! Добавление форматов невозможно. Нажмите любую клавишу для выхода");
                Console.ReadKey();
                return;
            }

            int sum = 0, err = 0, c = 0;

            foreach (string line in lines)
            {
                if (line.StartsWith("#")) continue;
                string[] data = line.Split(';')[0].Split(',');

                //все действия с форматами производим в книжной ориентации
                double heigthMm = double.Parse(data[0]);
                double widthMm = double.Parse(data[1]);
                string formatName = data[2];

                Console.WriteLine();
                Console.WriteLine("Формат: " + formatName + " " + widthMm.ToString("F0") + "x" + heigthMm.ToString("F0") + "(h). Проверяем наличие... ");


                System.Drawing.Printing.PaperSize winPaperSize = PrinterUtility.GetPaperSize(printerName, widthMm, heigthMm);
                if (winPaperSize != null)
                {
                    Console.WriteLine("Уже существует! Пропущен");
                    err++;
                }
                else
                {
                    PrinterUtility.AddFormat("PDFCreator",formatName, widthMm / 10, heigthMm / 10);
                    System.Threading.Thread.Sleep(100);
                    Console.WriteLine("Формат успешно добавлен");
                    c++;
                }
                sum++;
            }
            Console.WriteLine();
            Console.WriteLine("Успешно добавлено: " + c.ToString());
            Console.WriteLine("Ошибок: " + err.ToString());
            Console.WriteLine("Обработано форматов: " + sum.ToString());

            try
            {
                batchPrintSubkey.SetValue(registryKey, "true");
                //Microsoft.Win32.Registry.SetValue(registryPath, registryKey, "True", Microsoft.Win32.RegistryValueKind.String);
                Console.WriteLine("Добавлена пометка в реестр " + path);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Не удалось внести пометку в реестр: " + ex.Message);
            }

            if (!args.Contains("/c"))
            {
                Console.WriteLine("Нажмите любую клавишу для выхода");
                Console.ReadKey();
            }
        }
    }
}
