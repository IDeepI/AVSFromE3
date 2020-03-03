using E3Namespace;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace AVSFromE3
{
    class Program
    {
        static readonly DateTime TimeStart = DateTime.Now;
        static int currentLine = 0;
        const string LineBegin = "23 49";
        const string BeforImbase = "2F";
        const string LastLine = "3C 4E 4F 5F 49 4E 49 5F 53 45 43 54 49 4F 4E 3E";
        public static List<byte[]> listOfBytes = new List<byte[]>();
        public static List<byte[]> titleListOfBytes = new List<byte[]>(30);

        static void Main()
        {
            Debug.WriteLine(TimeStart);
            // Собираем данные устройств
            E3Devices.GetDevices();

            Debug.WriteLine(DateTime.Now - TimeStart);

            string filename = "d:\\SAVE\\DOC\\1test\\Test Лист 1.PE";
            // Пишем в файл
            WriteBinaryFile(filename, GeneratePE());
        }
        // Собираем byte массив
        static byte[] GeneratePE()
        {                       

            // Создаем массив данных устройств
            GenOrdinarLine(ref listOfBytes);

            // Линий в паспорте
            byte[] lineInPE = {(byte)(listOfBytes.Count)}; // 
            string lineInPasport = "11"; // + 1 если есть исполнение            

            // Константы
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("69 53 50 32 01 00 00"));    // 0            
            titleListOfBytes.Add(lineInPE);
            //titleListOfBytes.Add(GetByteArrayFromString("output.pe"));    // 2
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("00 00 00 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 23 49"));      // 3
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers(lineInPasport));    // 4   

            //titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("9F 00 C3 00 CB 00 D3 00 DA 00 DD 00 E6 00 E9 00 EB 00 ED 00 44 01 76 01 0F 84")); //6           
            //Title(6) = Title(6) & shiftHex(symbcnt(p))  Знаков в Обозначение|Наименование	|Выполнил и др.

            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("0F 84 1A 01 02 09 10 12 1B 08 0C 7B 7C 7A A9 A6 CC"));      // 7
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("43 3A 5C 44 6F 63 75 6D 65 6E 74 73 20 61 6E 64 20 53 65 74 74 69 6E 67 73 5C 61 6E 64 72 75 73 65 76 69 63 68 5C CC EE E8 20 E4 EE EA F3 EC E5 ED F2 FB 5C 43 41 44 65 6C 65 63 74 72 6F 20 CF F0 E8 EC E5 F0 5C D9 D1 CD 20 30 36 32 2D 30 36 36 5C D9 D1 CD 30 36 33 2E 70 65 69 00"));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("5C 5C 53 51 4C 4A 55 4E 49 4F 52 5C 49 4D 5C 41 56 53 35 5C 41 56 53 35 5F 50 45 33 2E 49 4E 49 00"));    // 8

            //titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("5C 5C 53 51 4C 4A 55 4E 49 4F 52 5C 49 4D 5C 41 56 53 35 5C 41 56 53 35 5F 50 45 33 2E 49 4E 49 00"));    // 9

            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("6F 75 74 70 75 74 2E 50 45 00 "));    // 11
            titleListOfBytes.Add(GetByteArrayFromString(E3Devices.prjNum));
            titleListOfBytes.Add(GetByteArrayFromString(E3Devices.placename));
            titleListOfBytes.Add(GetByteArrayFromString(E3Devices.executor));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("D4 E5 E4 EE F0 EE E2 00 "));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("C8 E2 E0 ED EE E2 00"));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("C0 34 00 "));
            titleListOfBytes.Add(GetByteArrayFromString(DateTime.Now.ToShortDateString()));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("EA E3 00 "));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("30 00"));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("31 00"));    // 21
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("31 38 00"));    // 22
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("D4 E0 E9 EB 20 F1 EE F5 F0 E0 ED E5 ED 3A 20 49 54 30 30 30 39 36 20 30 37 2E 30 36 2E 32 30 31 39 20 31 35 3A 34 36 3A 33 30 20 5C 5C 43 53 31 34 5C 49 4D 5C 41 56 53 35 5C 41 56 53 35 2E 45 58 45 20 41 56 53 20 35 2E 37 2E 38 32 75 20 4E 65 74 77 6F 72 6B 00 "));     // 23
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("41 56 53 20 35 2E 37 2E 38 32 75 20 4E 65 74 77 6F 72 6B 20 45 78 74 65 6E 64 65 64 20 3A 20 30 35 2E 31 32 2E 32 30 31 38 20 31 36 3A 32 35 3A 30 30 00 "));    // 24
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("30"));    // 25

            //Call GenTitleSix(25)

            // titleListOfBytes.Add(GetByteArrayFromStringOfNumbers(""));
            titleListOfBytes.Add(GetByteArrayFromStringOfNumbers("")); // 27


            titleListOfBytes.Insert(4, GenTitleCnt()); // 6


            for (int j = titleListOfBytes.Count-1; j >= 0; j--)
            {
                listOfBytes.Insert(0, titleListOfBytes[j]);
            }

            listOfBytes.Add(GetByteArrayFromStringOfNumbers("3C 4E 4F 5F 49 4E 49 5F 53 45 43 54 49 4F 4E 3E"));

            /*          
                  Title(2) = HexD(line - 1)//' Колличество строк
                  Title(7) = "0F841A01020910121B080C7B7C7AA9A6CC"
                  Title(22) = "313800"
                  Title(25) = "3000"
                  Erase symbcnt

              Totalsymbcnt = 0
              For p = 8 to 30
                  If Title(p) <> "" Then
                      symbcnttmp = Len(Title(p)) / 2
                      Totalsymbcnt = Totalsymbcnt + symbcnttmp
                      symbcnt(p) = Totalsymbcnt

                      IF p<TitleCnt Then
                          Title(6) = Title(6) & shiftHex(symbcnt(p)) '"9F00C300CB00D300DA00DD00E600E900EB00ED00440176010F84"
                      End If
                  End If
              Next
             Title(4) = shiftHex(Totalsymbcnt)     

              Title(3) = "00000020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202349" & lineInPasport
              Title(5) = HexD(IspLine) & "000000"'"00000000" '5D007E008800" ' 01 02 кол исп в групповая спецификация
              */

            // Собираем все в один масссив 
            for (int i = 1; i < listOfBytes.Count; i++)
            {
                try
                {
                    listOfBytes[0] = listOfBytes[0].Concat(listOfBytes[i]).ToArray();
                }
                catch
                {
                   // Debug.WriteLine($"Нет {i}-го члена массива.");
                }
            }

            return listOfBytes[0];
        }

        static byte[] GenTitleCnt()
        {
            byte[] byteArray = new byte[0];
            int totalByte = 0;
            for (int i = 5; i< titleListOfBytes.Count(); i++)
            {
                totalByte += titleListOfBytes[i].Length;
                if (i < 21)
                {
                    byteArray = byteArray.Concat(GetByteArrayFromInt(totalByte)).ToArray();
                }
                
            }
            
            byteArray = GetByteArrayFromInt(totalByte)
                .Concat(GetByteArrayFromStringOfNumbers("00 00 00 00"))                
                .Concat(byteArray)
                .ToArray();

            return byteArray;
        }


        static void GenOrdinarLine(ref List<byte[]> arrLine)
        {
            foreach (E3Device device in E3Devices.devListSorted)
            {/*
                Debug.WriteLine($"Свойство 0 - {device.Properties[0]}");    // Имя изделия
                Debug.WriteLine($"Свойство 1 - {device.Properties[1]}");    // Код ImBase
                Debug.WriteLine($"Свойство 2 - {device.Properties[2]}");    // Наименование
                Debug.WriteLine($"Свойство 3 - {device.Properties[3]}");    // Примечание
                Debug.WriteLine($"Свойство 4 - {device.Properties[4]}");    // Исполнение*/
               // Debug.WriteLine($"Свойства - {device.ToString()}");    // 

                int ispsymbCnt = 0;
                string razdel = " ";
                currentLine = currentLine + 1;
                int symbInProp = 0;
                int[] symbCnt = new int[9];
                byte[][] propertiesByteArray = new byte[7][];

                // Преобразуем свойства в byte[]
                int n = 0;
                foreach (string propertie in device.Properties)
                {
                    propertiesByteArray[n++] = GetByteArrayFromString(propertie);
                }

                //  Считаем колличество символов в полях                       

                for (int m = 0; m < propertiesByteArray.Count(); m++)
                {
                    symbInProp = propertiesByteArray[m].Length;
                    symbCnt[m] = (m == 0) ? symbInProp : symbCnt[m - 1] + symbInProp; // Учитываем разделительные 00
                }

                arrLine.Add(
                    GetByteArrayFromStringOfNumbers(LineBegin)
                    .Concat(GenStr(symbCnt))
                    .Concat(GetByteArrayFromStringOfNumbers(BeforImbase))
                    .Concat(GetByteArrayFromString(device.Properties[0]))
                    .Concat(GetByteArrayFromString(device.Properties[1]))
                    .Concat(GetByteArrayFromString(device.Properties[2]))
                    .Concat(GetByteArrayFromString(device.Properties[3]))
                    .Concat(GetByteArrayFromString(device.Properties[4]))
                    .Concat(GetByteArrayFromString(device.Properties[5]))
                    .Concat(GetByteArrayFromString(device.Properties[6]))                                      
                    .ToArray()
                );
            }
        }

        static byte[] GenStr(int[] symbCnt)
        {
            return GetByteArrayFromStringOfNumbers("07")
            .Concat(GetByteArrayFromInt(symbCnt[6]))
            .Concat(GetByteArrayFromStringOfNumbers("00 00 00 00"))
            .Concat(GetByteArrayFromInt(symbCnt[0]))
            .Concat(GetByteArrayFromInt(symbCnt[1]))
            .Concat(GetByteArrayFromInt(symbCnt[2]))
            .Concat(GetByteArrayFromInt(symbCnt[3]))
            .Concat(GetByteArrayFromInt(symbCnt[4]))
            .Concat(GetByteArrayFromInt(symbCnt[5]))
            .Concat(GetByteArrayFromStringOfNumbers("08 0A 20 05 06 07"))
            .ToArray();
        }



        static byte[] GetByteArrayFromString(string res)
        {
            if (res == "")
            {
                res = " ";
            }
            byte[] byteArray = System.Text.Encoding.GetEncoding(1251).GetBytes(res);
            Array.Resize(ref byteArray, byteArray.Length + 1);
            byteArray[byteArray.Length - 1] = 0;

            return byteArray;
        }
        static byte[] GetByteArrayFromInt(int res)
        {
            string hexValue;
            byte[] bytes = BitConverter.GetBytes(res);
           // Debug.WriteLine(BitConverter.ToString(bytes));
            hexValue = BitConverter.ToString(bytes).Substring(0,5).Replace("-", " ");
          //  Debug.WriteLine(hexValue);
            return GetByteArrayFromStringOfNumbers(hexValue);
        }

        static byte[] GetByteArrayFromStringOfNumbers(string dataString)
        {
            string hexToByte;

            dataString = dataString.Trim();
            string[] hexValuesSplit = dataString.Split(' ');

            byte[] byteValue = new byte[(hexValuesSplit.Length)];
            int i = 0;
            
            if (dataString != "")
            {
                
                foreach (string hex in hexValuesSplit)
                {                    
                    if (hex.Length % 2 != 0)
                    {
                        hexToByte = "0" + hex;
                    }else
                    {
                        hexToByte = hex;
                    }
                    // Convert the number expressed in base-16 to an integer.   
                    byteValue[i++] = (byte)Convert.ToInt32(hexToByte, 16);
                }
                
            }
          
            return byteValue;
        }

        static void WriteBinaryFile(string filename, byte[] byteValue)
        {

            // создаем файл
            try
            {
                //  System.Text.Encoding.GetEncoding(1251) //  "Windows — 1251" ("Windows — 1251") 
                // byte[] messageByte = Encoding.ASCII.GetBytes("Here is some data.");
                FileStream fstr = new FileStream(filename, FileMode.Create);
                fstr.Write(byteValue, 0, byteValue.Length);
                fstr.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            Console.ReadLine();
        }
    }
}


