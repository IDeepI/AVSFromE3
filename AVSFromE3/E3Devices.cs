using ConnectToE3;
using e3;
using ImBaseExtensionDLL;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace E3Namespace
{
    public static class StringExtensions
    {
        public static string Right(this string str, int length)
        {
            return str.Substring(str.Length - length, length);
        }
    }
    public static class E3Devices
    {
        public static e3Application App; // объект приложения
        public static e3ExternalDocument Ex; // объект приложения
        public static e3Job Prj; // объект проекта
        public static e3Symbol Sym;	// объект 
        public static e3Device Dev;	// объект 
        public static e3Component Cmp;	// объект 
        public static e3Pin Pin;    // объект 
        public static e3DbeApplication e3Dbe;// объект приложения
        public static List<E3Device> devList = new List<E3Device>();
        public static List<E3Device> devListSorted = new List<E3Device>();

        ///////////////////////////////////////
        public static string prjNum;
        public static string placename;
        public static string executor;
        public static string AVSName;
        public static string AVSFileNew;
        /////////////////////////////////////////


        public static void GetE3App()
        {
            // Подключаем E3
            App = AppConnect.ToE3();
            App?.PutInfo(0, "GetE3App!");
        }
        public static void GetE3App(string filePath)
        {
            // Подключаем E3
            App = AppConnect.ToE3(filePath);
            App?.PutInfo(0, $"GetE3App for {filePath}!");
        }

        public static void LoadAvsToE3()
        {
            App?.PutInfo(0, "Load AVS to E3!");
            Ex = (e3ExternalDocument)Prj.CreateExternalDocumentObject();
            // Объекты массивов Id
            Object exIds = new Object();
            string pename = "";
            Prj.GetExternalDocumentIds(ref exIds);
            // Удаляем существующий PE
            try
            {
                foreach (var exId in (Array)exIds)
                {
                    if (exId != null)
                    {
                        Ex.SetId((int)exId);
                        pename = Ex.GetName();
                        if (pename.Right(3) == ".PE")
                        {
                            Ex.Delete();
                            App.PutInfo(0, $"Deleted {pename}");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                App.PutInfo(0, $"Error with {pename}");
                App.PutInfo(0, $"First exception caught. {e}");
            }

            Ex.Create(0, AVSName, AVSFileNew); //' Вставить новый AVS

            App?.PutInfo(0, $"File inserted - {AVSFileNew}");
        }


        public static void GetDevices()
        {
            // Объекты массивов Id
            Object DevIds = new Object();
            //Object textIds = new Object();
            string placedName;


            App?.PutInfo(0, "Starting GetDevices!");
            Prj = (e3Job)App?.CreateJobObject();
            Sym = (e3Symbol)Prj.CreateSymbolObject();
            Dev = (e3Device)Prj.CreateDeviceObject();
            Cmp = (e3Component)Prj.CreateComponentObject();
            Pin = (e3Pin)Prj.CreatePinObject();

            // Получаем массив Id устройств
            int nd = Prj.GetAllDeviceIds(ref DevIds);
            // Переменные среды
            int PinCnt;
            Object PinIds = new Object();
            string imbaseKEY;
            bool DoItAgain = false;

            ///////////////////////////////////////
            prjNum = Prj.GetAttributeValue("NUMPROJ") + " ПЭ3";
            placename = Prj.GetAttributeValue("OBOZNACHENIE") + " " + Prj.GetAttributeValue("Uslovn_obozn") + " " + Prj.GetAttributeValue("Name_of_schemes") + ". Перечень элементов";
            executor = Prj.GetAttributeValue("vyp");


            if (Prj.GetName().Length > 18)
            {
                AVSName = Prj.GetName().Substring(0, 18) + ".PE";
            }
            else
            {
                AVSName = Prj.GetName() + ".PE";
            }
            AVSFileNew = Prj.GetPath() + AVSName; //'Имя файла AVS(без пробелов) 
            /////////////////////////////////////////



            foreach (var devId in (Array)DevIds)
            {
                if (devId != null)
                {
                    Dev.SetId((int)devId);
                    Cmp.SetId((int)devId);

                    if (Dev.IsWireGroup() == 0 && Dev.IsCable() == 0 && Dev.GetComponentName() != "" && Dev.IsAssembly() == 0)
                    {
                        bool isTerminal = (Dev.IsTerminal() == 1 && Dev.IsTerminalBlock() != 1);
                        string pinName = null;
                        // Позиционное обозначение
                        placedName = Dev.GetName();
                        // Если клемма
                        if (isTerminal)
                        {
                            PinCnt = Dev.GetPinIds(ref PinIds);
                            if (PinCnt > 0) // если выводов нет, то и клемма не попадет в отчет...
                            {
                                foreach (var pinId in (Array)PinIds)
                                {
                                    if (pinId == null)
                                        continue;
                                    Pin.SetId((int)pinId);
                                    if (Pin.GetName() != "")
                                    {
                                        pinName = Pin.GetName();
                                    }
                                    break;
                                }
                            }
                        }

                        imbaseKEY = Cmp.GetAttributeValue("Imbase_KEY");
                        if (imbaseKEY.Length != "i609010608038500008C".Length && imbaseKEY.Length != 1)
                        {
                            Prj.JumpToID((int)devId);
                            e3Dbe = new e3DbeApplication(); // объект приложения                          
                            e3Dbe.EditComponent(Dev.GetComponentName(), Dev.GetComponentVersion());
                            //        msgbox "Не корректный ключ Imbase - " & ImbaseKEY & " - изделия - " & dev.getname
                            Console.WriteLine($"Не корректный ключ Imbase - {imbaseKEY} - изделия - {Dev.GetName()}");
                            Console.WriteLine("Нажмите 'Enter' после корректировки изделия");
                            Console.ReadLine();
                            //'обновить																
                            Prj.UpdateComponent(Cmp.GetName(), 0);
                            imbaseKEY = Cmp.GetAttributeValue("Imbase_KEY");
                            DoItAgain = true;
                        }
                        // devList.Add(new E3Device(placedName, new string[] {Dev.GetComponentName(), imbaseKEY, Cmp.GetAttributeValue("Description"), Dev.GetAttributeValue("Primechanie"), Dev.GetAttributeValue("Исполнение")}, pinName));


                        if (imbaseKEY.Length == "i609010608038500008C".Length)
                        {

                            if (isTerminal)
                            {
                                devList.Add(new E3Device(placedName, new string[] { "I" + imbaseKEY, "6", placedName + ":" + pinName, ImBaseEx.GetFullDesignation(imbaseKEY), "1", Dev.GetAttributeValue("Primechanie"), "1" }, pinName));
                            }
                            else
                            {
                                devList.Add(new E3Device(placedName, new string[] { "I" + imbaseKEY, "6", placedName, ImBaseEx.GetFullDesignation(imbaseKEY), "1", Dev.GetAttributeValue("Primechanie"), "1" }, pinName));
                            }
                        }
                        else
                        {
                            if (isTerminal)
                            {
                                devList.Add(new E3Device(placedName, new string[] { "Нет ключа Imbase", "5", placedName + ":" + pinName, Cmp.GetAttributeValue("Description"), "1", Dev.GetAttributeValue("Primechanie"), "1" }, pinName));
                            }
                            else
                            {
                                devList.Add(new E3Device(placedName, new string[] { "Нет ключа Imbase", "5", placedName, Cmp.GetAttributeValue("Description"), "1", Dev.GetAttributeValue("Primechanie"), "1" }, pinName));
                            }
                        }

                    }
                    if (DoItAgain)
                    {
                        // Думаю можно продолжать без выхода
                        // Environment.Exit(0);
                    }
                }
            }
            // Сортировка устройств
            devList.Sort(delegate (E3Device x, E3Device y)
            {
                if (x.IsTerminal() && x.Name == y.Name)
                {
                    if (x.SortPinValue() == y.SortPinValue())
                    {
                        return x.PinName.CompareTo(y.PinName);
                    }
                    else
                    {
                        return x.SortPinValue().CompareTo(y.SortPinValue());
                    }
                }
                else
                {
                    if (GetLetter(x.Name) == GetLetter(y.Name))
                    {
                        return x.SortNameValue().CompareTo(y.SortNameValue());
                    }
                    else
                    {
                        return x.Name.CompareTo(y.Name);
                    }
                }
            });

            foreach (E3Device dev in devList)
            {
                //   Console.WriteLine(dev);                
                SortDevList(dev);
            }

            Debug.Flush();
        }

        static string GetLetter(string textToParse)
        {
            Regex laterEx = new Regex(@"(?i)([a-z]+)(?:\d+)");

            MatchCollection matches = laterEx.Matches(textToParse);
            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    // Debug.WriteLine(match.Groups[1].Value);
                    return match.Groups[1].Value;
                }
            }
            return "";
        }

        static int GetNumber(string textToParse)
        {
            Regex numberEx = new Regex(@"(?i)(?:[a-z]+)(\d+)");

            MatchCollection matches = numberEx.Matches(textToParse);
            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    // Debug.WriteLine(match.Groups[1].Value);
                    return Int32.Parse(match.Groups[1].Value);
                }
            }
            return 1;
        }

        static int GetPinNumber(string textToParse)
        {
            Regex numberEx = new Regex(@"(?::)(\d+)");

            MatchCollection matches = numberEx.Matches(textToParse);
            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    // Debug.WriteLine(match.Groups[1].Value);
                    return Int32.Parse(match.Groups[1].Value);
                }
            }
            return 1;
        }

        static bool CheckComma(string textToParse)
        {
            Regex numberEx = new Regex(@"(,\w+:*\w*)$");

            MatchCollection matches = numberEx.Matches(textToParse);
            if (matches.Count > 0)
            {
                return true;
            }
            return false;
        }

        static bool CheckColon(string textToParse)
        {
            Regex numberEx = new Regex(@"(\.\.\w+:*\w*)$");

            MatchCollection matches = numberEx.Matches(textToParse);
            if (matches.Count > 0)
            {
                return true;
            }
            return false;
        }

        static int GetPreLastNumber(string textToParse)
        {
            Regex numberEx = new Regex(@"(?i)(\d+)$");

            MatchCollection matches = numberEx.Matches(textToParse);
            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    // Debug.WriteLine(match.Groups[1].Value);
                    return Int32.Parse(match.Groups[1].Value);
                }
            }
            return 1;
        }



        static string GetShorter(string textToParse)
        {
            string patern = (@"(?i)(\.\.\w+:*\w*)$|(,\w+:*\w*)$");
            return Regex.Replace(textToParse, patern, "");
        }

        static int GetNumberOnly(string textToParse)
        {
            string patern = (@"(?i)(\D+)");
            return Int32.Parse(Regex.Replace(textToParse, patern, ""));
        }




        static void SortDevList(E3Device device)
        {
            string placeName;

            // Проверям есть ли устройство с таким описанием
            E3Device tmpDev = devListSorted.Find(x => x.Properties[3].Contains(device.Properties[3])); // Designation
            if (tmpDev != null)
            {
                placeName = tmpDev.Properties[2];

                // Если есть добавляем к существующему устройству
                if (device.IsTerminal())
                {
                    // Если имена совпадают
                    if (tmpDev.Name == device.Name)
                    {
                        // Номера отличаются на 1
                        if (GetNumberOnly(tmpDev.PinName) == GetNumberOnly(device.PinName) - 1)
                        {
                            // две точки уже есть
                            if (CheckColon(placeName))
                            {
                                placeName = GetShorter(placeName) + ".." + device.Name + ":" + device.PinName; // placedName + ":" + pinName
                            }
                            // Запятая уже есть и число 3-е подряд
                            else if (CheckComma(placeName) && GetPreLastNumber(GetShorter(placeName)) == (GetNumberOnly(device.PinName) - 2))
                            {
                                placeName = GetShorter(placeName) + ".." + device.Name + ":" + device.PinName; // placedName + ":" + pinName
                            }
                            else
                            {
                                placeName = placeName + "," + device.Name + ":" + device.PinName; // placedName + ":" + pinName
                            }

                        }
                        else
                        {
                            placeName = placeName + "," + device.Name + ":" + device.PinName; // placedName + ":" + pinName
                        }

                    }
                    else
                    {
                        placeName += "," + device.Name + ":" + device.Name + ":" + device.PinName;// placedName + ":" + pinName
                    }
                }
                else
                {
                    // Если имена совпадают
                    if (GetLetter(tmpDev.Name) == GetLetter(device.Name))
                    {
                        // Номера отличаются на 1
                        if (GetNumberOnly(tmpDev.Name) == GetNumberOnly(device.Name) - 1)
                        {
                            // две точки уже есть
                            if (CheckColon(placeName))
                            {
                                placeName = GetShorter(placeName) + ".." + (device.Name); // placedName + ":" + pinName
                            }
                            // Запятая уже есть и число 3-е подряд
                            else if (CheckComma(placeName) && GetPreLastNumber(GetShorter(placeName)) == (GetNumberOnly(device.Name) - 2))
                            {
                                placeName = GetShorter(placeName) + ".." + (device.Name); // placedName + ":" + pinName
                            }
                            else
                            {
                                placeName = placeName + "," + (device.Name); // placedName + ":" + pinName
                            }
                        }
                        else
                        {
                            placeName += "," + (device.Name); // placedName 
                        }
                    }
                    else
                    {
                        placeName += "," + device.Name;// placedName 
                    }
                }
                tmpDev.Properties[2] = placeName;
                tmpDev.PinName = device.PinName; // Для последующей сортировки
                tmpDev.Name = device.Name; // Для последующей сортировки

                tmpDev.Properties[4] = (Int32.Parse(tmpDev.Properties[4]) + 1).ToString(); // Count
                tmpDev.Properties[6] = (Int32.Parse(tmpDev.Properties[6]) + 1).ToString(); // Count
            }
            // Если нет то добавляем
            else
            {
                //  Добавляем новое устройство
                devListSorted.Add(device);
            }
        }
    }

    public class E3Device
    {
        public string Name { get; set; }
        public string Count { get; set; }
        public string PinName { get; set; }
        public string[] Properties { get; set; }
        public bool IsTerminal() => PinName != null;
        public string PlacedName() => IsTerminal() ? Name + ":" + PinName : Name;
        public int SortPinValue() => Convert.ToInt32(Regex.Replace(PinName, @"\D", "", RegexOptions.IgnoreCase));
        public int SortNameValue() => Convert.ToInt32(Regex.Replace(Name, @"\D", "", RegexOptions.IgnoreCase));
        public override string ToString() => Name + "|" + string.Join("|", Properties);

        public E3Device() : this("Неизвестно") { }
        public E3Device(string name)
        {
            this.Name = name;
            this.Count = "1";

        }
        public E3Device(string name, string[] properties)
        {
            this.Name = name;
            this.Properties = properties;
        }
        public E3Device(string name, string[] properties, string pinName)
        {
            this.Name = name;
            this.Properties = properties;
            this.PinName = pinName;
        }
    }
}

