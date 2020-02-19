using ConnectToE3;
using e3;
using ImBaseExtensionDLL;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace E3Namespace
{
    public static class E3Devices
    {
        public static e3Application App; // объект приложения
        public static e3Job Prj; // объект проекта
        public static e3Symbol Sym;	// объект 
        public static e3Device Dev;	// объект 
        public static e3Component Cmp;	// объект 
        public static e3Pin Pin;    // объект 
        public static e3DbeApplication e3Dbe;// объект приложения
        public static List<E3Device> devList = new List<E3Device>();

        public static void GetDevices()
        {
            // Объекты массивов Id
            Object DevIds = new Object();
            //Object textIds = new Object();
            string placedName;

            // Подключаем E3
            App = AppConnect.ToE3();
            App?.PutInfo(0, "Starting E3Devices!");
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
                            devList.Add(new E3Device(placedName, new string[] { Dev.GetComponentName(), imbaseKEY, ImBaseEx.GetFullDesignation(imbaseKEY), Dev.GetAttributeValue("Primechanie"), Dev.GetAttributeValue("Исполнение") }, pinName));
                        }
                        else
                        {
                            devList.Add(new E3Device(placedName, new string[] { Dev.GetComponentName(), imbaseKEY, Cmp.GetAttributeValue("Description"), Dev.GetAttributeValue("Primechanie"), Dev.GetAttributeValue("Исполнение") }, pinName));
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
                    if (x.SortValue() == y.SortValue())
                    {
                        return x.PinName.CompareTo(y.PinName);
                    }
                    else
                    {
                        return x.SortValue().CompareTo(y.SortValue());
                    }

                }
                else
                {
                    return x.Name.CompareTo(y.Name);
                }
            });

            foreach (E3Device dev in devList)
            {
             //   Console.WriteLine(dev);
            }
            Debug.Flush();
        }
    }

    public class E3Device
    {
        public string Name { get; set; }
        public string PinName { get; set; }
        public string[] Properties { get; set; }
        public bool IsTerminal() => PinName != null;
        public string PlacedName() => IsTerminal() ? Name + ":" + PinName : Name;
        public int SortValue() => Convert.ToInt32(Regex.Replace(PinName, @"\D", "", RegexOptions.IgnoreCase));
        public override string ToString() => Name + "|" + PinName + "|" + string.Join("|", Properties);

        public E3Device() : this("Неизвестно") { }
        public E3Device(string name)
        {
            this.Name = name;
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

