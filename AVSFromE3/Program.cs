using E3Namespace;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace AVSFromE3
{
    class Program
    {
        static readonly DateTime TimeStart = DateTime.Now;

        static void Main()
        {
            Debug.WriteLine(TimeStart);

            E3Devices.GetDevices();

           // Debug.WriteLine(E3Devices.devList[1]);
            Debug.WriteLine(DateTime.Now - TimeStart);

            string filename = "d:\\SAVE\\DOC\\1test\\Test ВЕАШ.674722.472 - 09.PE";
            WriteBinaryFile(filename, GeneratePE());
        }

        static byte[] GeneratePE()
        {
            const string LineBegin = "2349";
            const string BeforImbase = "2F";
            const string LastLine = "3C4E4F5F494E495F53454354494F4E3E";


            byte[] LineInPasport = { 11 }; // + 1 если есть исполнение

            byte[][] arrTitle = new byte[30][];
            arrTitle[0] = GetByteArrayFromStringOfNumbers("69 53 50 32 01 00 00");
            arrTitle[2] = GetByteArrayFromString("output.pe");
            arrTitle[3] = GetByteArrayFromStringOfNumbers("00 00 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 23 49");
            arrTitle[4] = LineInPasport;



            arrTitle[8] = GetByteArrayFromStringOfNumbers("43 3A 5C 44 6F 63 75 6D 65 6E 74 73 20 61 6E 64 20 53 65 74 74 69 6E 67 73 5C 61 6E 64 72 75 73 65 76 69 63 68 5C CC EE E8 20 E4 EE EA F3 EC E5 ED F2 FB 5C 43 41 44 65 6C 65 63 74 72 6F 20 CF F0 E8 EC E5 F0 5C D9 D1 CD 20 30 36 32 2D 30 36 36 5C D9 D1 CD 30 36 33 2E 70 65 69 00 5C 5C 53 51 4C 4A 55 4E 49 4F 52 5C 49 4D 5C 41 56 53 35 5C 41 56 53 35 5F 50 45 33 2E 49 4E 49 00");
            arrTitle[9] = GetByteArrayFromStringOfNumbers("5C 5C 53 51 4C 4A 55 4E 49 4F 52 5C 49 4D 5C 41 56 53 35 5C 41 56 53 35 5F 50 45 33 2E 49 4E 49 00");

            arrTitle[11] = GetByteArrayFromStringOfNumbers("6F 75 74 70 75 74 2E 50 45 00 ");
            arrTitle[12] = GetByteArrayFromString(E3Devices.prjNum);
            arrTitle[13] = GetByteArrayFromString(E3Devices.placename);
            arrTitle[14] = GetByteArrayFromString(E3Devices.executor);
            arrTitle[15] = GetByteArrayFromStringOfNumbers("D4 E5 E4 EE F0 EE E2 00 ");
            arrTitle[16] = GetByteArrayFromStringOfNumbers("C8 E2 E0 ED EE E2 00");
            arrTitle[17] = GetByteArrayFromStringOfNumbers("C0 34 00 ");
            arrTitle[18] = GetByteArrayFromString(DateTime.Now.ToShortDateString());
            arrTitle[19] = GetByteArrayFromStringOfNumbers("EA E3 00 ");
            arrTitle[20] = GetByteArrayFromStringOfNumbers("30 00");
            arrTitle[21] = GetByteArrayFromStringOfNumbers("31 00");

            arrTitle[23] = GetByteArrayFromStringOfNumbers("D4 E0 E9 EB 20 F1 EE F5 F0 E0 ED E5 ED 3A 20 49 54 30 30 30 39 36 20 30 37 2E 30 36 2E 32 30 31 39 20 31 35 3A 34 36 3A 33 30 20 5C 5C 43 53 31 34 5C 49 4D 5C 41 56 53 35 5C 41 56 53 35 2E 45 58 45 20 41 56 53 20 35 2E 37 2E 38 32 75 20 4E 65 74 77 6F 72 6B 00 ");
            arrTitle[24] = GetByteArrayFromStringOfNumbers("41 56 53 20 35 2E 37 2E 38 32 75 20 4E 65 74 77 6F 72 6B 20 45 78 74 65 6E 64 65 64 20 3A 20 30 35 2E 31 32 2E 32 30 31 38 20 31 36 3A 32 35 3A 30 30 00 ");
            // arrTitle[27] = GetByteArrayFromStringOfNumbers("");

            // Создаем массив данных hex - значений в string формате

            GenOrdinarLine(ref arrTitle);

            GenStrArr(ref arrTitle);


            for (int i = 1; i < arrTitle.Length; i++)
            {
                try
                {
                    arrTitle[0] = arrTitle[0].Concat(arrTitle[i]).ToArray();
                }
                catch
                {
                    Debug.WriteLine($"Нет {i}-го члена массива.");
                }
            }

            return arrTitle[0];
        }

        static void GenOrdinarLine(ref byte[][] arrLine)
        {
            foreach(E3Device device in E3Devices.devList)
            {
                Debug.WriteLine($"Свойство 0 - {device.Properties[0]}");    // Имя изделия
                Debug.WriteLine($"Свойство 1 - {device.Properties[1]}");    // Код ImBase
                Debug.WriteLine($"Свойство 2 - {device.Properties[2]}");    // Наименование
                Debug.WriteLine($"Свойство 3 - {device.Properties[3]}");    // Примечание
                Debug.WriteLine($"Свойство 4 - {device.Properties[4]}");    // Исполнение
                 



            }
        }

        static void GenStrArr(ref byte[][] arrStr)
        {

        }

        /*       Sub GenOrdinarLine()

           ' строки без исполнений
           For line = 1 to all	

               if Len(PEArr(7, line - 1)) < 2 Then
                   CurrentLine = CurrentLine + 1

                   Erase symbcnt

                   Erase Data

                   Totalsymbcnt = 0  


                   For i = 0 to 7
                       if i< 7 Then

                           Data(i)	= String2Asc(PEArr(i, line-1)) & "00"	' Imbase code
                       ElseIf Len(PEArr(7, line - 1)) > 1  Then
                           Data(i) = String2Asc(PEArr(i, line-1)) & "00"	' Imbase code
                       End If

                       symbcnt(i) =Len(Data(i))/2		

                       Totalsymbcnt = Totalsymbcnt + symbcnt(i)
                       'msgbox Totalsymbcnt,," Totalsymbcnt"
                       If i > 0 Then
                           symbcnt(i) = symbcnt(i) + symbcnt(i-1)

                       End If


                   Next
                   Redim Preserve StrArr(CurrentLine + 1)

                   StrArr(CurrentLine) = LineBegin & HashTxt(IspSymbCnt) & BeforImbase & Data(0) & Data(1)  & Data(2) & Data(3) & Data(4) & Data(5) & Data(7) & Data(6) 'записываем строки ПЭ в массив
               End IF

           Next
       End Sub

       Sub GenStrArr()	
       

           ' Строки с исполнениями
           ' Словарь для исполнений
           RazdelIsp = "23560121000000000005CFE5F0E5ECE5EDEDFBE520E4E0EDEDFBE520E4EBFF20E8F1EFEEEBEDE5EDE8E900"
           Dim DictIsp, Line

           Set DictIsp = CreateObject("Scripting.Dictionary")
               DictIsp.CompareMode = vbTextCompare

           For line = 1 to all		
               if Len(PEArr(7, line - 1)) > 1 Then				
                   IspData = PEArr(7, line-1) ' Исполнение			
                   'msgbox IspData & "-" & line & "-" & all
                   If Not DictIsp.Exists(IspData) Then 
                       IspLine = IspLine + 1
                       DictIsp.Add IspData, 1
                       Call GenIspLine(IspData)			

                   End IF

               End If	
           Next

           Redim Preserve StrArr(line + 2)

           StrArr(line+1) =  LastLine'записываем строки ПЭ в массив
           If IspLine > 0 Then
               LineInPasport = 11 + 1 ' + 1 если есть исполнение
               Title(2) = HexD(line + IspLine)' Колличество строк
               Title(7) =  "0F8401020910121B080C7B7C7A1AA9A606CC" 
               Title(21) = "313800"' ? "313800"'"3600" 
               Title(22) = "7465737420E8F1EFEEEBEDE5EDE8E92E504500" ' filename.pe
               Title(25) = "3100" ' "3000"
               Title(26) = "3100"
               Call  GenTitleSix(26)
               'msgbox 1
           Else
               LineInPasport = 11
               Title(2) = HexD(line - 1)' Колличество строк
               Title(7) =  "0F841A01020910121B080C7B7C7AA9A6CC" 		
               Title(22) = "313800"
               Title(25) = "3000"
               Call  GenTitleSix(25)
               'msgbox 2
           End If
           Title(3) ="00000020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202349" & LineInPasport
           Title(5) = HexD(IspLine) & "000000"'"00000000" '5D007E008800" ' 01 02 кол исп в групповая спецификация
           'msgbox  Title2,,"Title2"
           For t = 1 to UBound(Title)
               StrArr(0) = StrArr(0) + Title(t)
           Next 	
       End Sub

       */


        static byte[] GetByteArrayFromString(string res)
        {
            byte[] ByteArray = System.Text.Encoding.GetEncoding(1251).GetBytes(res);
            Array.Resize(ref ByteArray, ByteArray.Length + 1);
            ByteArray[ByteArray.Length - 1] = 0;
            return ByteArray;
        }

        static byte[] GetByteArrayFromStringOfNumbers(string dataString)
        {
            dataString = dataString.Trim();

            byte[] byteValue = new byte[(dataString.Length + 1) / 3];
            int i = 0;

            string[] hexValuesSplit = dataString.Split(' ');
            foreach (string hex in hexValuesSplit)
            {
                // Convert the number expressed in base-16 to an integer.    
                byteValue[i++] = (byte)Convert.ToInt32(hex, 16);
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


