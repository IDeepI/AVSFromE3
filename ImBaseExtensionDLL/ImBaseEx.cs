using ImBase;
using System;
using System.Diagnostics;

namespace ImBaseExtensionDLL
{
    public static class ImBaseEx
    {
        public static ImbaseApplicationClass ImApplication;
        public static ImDataBaseClass ImDataBase;

        public static void CloseImbaseConnection()
        {
            ImApplication = null;
            ImDataBase = null;
        }
        // Полное обозначение
        public static string GetFullDesignation(string ImCode, ref string codeAxapta)
        {
            // Подключаем базу
            ImDataBase = new ImDataBaseClass();
            while (ImDataBase.Ready() != 1)
            {
                System.Threading.Thread.Sleep(10);
            }
            // Получаем данные по клюючу ImBase
            ImDataBase.GetKeyInfo(ImCode, out string TableRecord, out string CatalogRecord, out string KeysList);
            //      Debug.WriteLine(TableRecord);
            // Debug.WriteLine(CatalogRecord);
            // Debug.WriteLine(KeysList);

            // Возвращаем Код Axapta  "Код Axapta=757478.1336",ПРИМЕЧАНИЕ=
            int startIndex = TableRecord.IndexOf("\"Код Axapta=") + 12;
            int endIndex = TableRecord.IndexOf("\",", startIndex);
            if (endIndex - startIndex > 0)
            {               
                codeAxapta = TableRecord.Substring(startIndex, endIndex - startIndex);
            }
            // Возвращаем ПОЛНОЕ ОБОЗНАЧЕНИЕ
            startIndex = CatalogRecord.IndexOf("\"ПОЛНОЕ ОБОЗНАЧЕНИЕ=") + 20;
            endIndex = CatalogRecord.IndexOf("\",КЛАСС=", startIndex);
            return CatalogRecord.Substring(startIndex, endIndex - startIndex);
        }


        public static bool CheckImbaseConnection()
        {
            ImApplication = new ImbaseApplicationClass();
            bool Done = false;
            try
            {
                // Создаем новое подключение, если указатель нулевой
                //
                //if (ImApplication == null) ImApplication = CoImbaseApplication.Create();
                // Проверяем состояние сервера
                while (!Done)
                { // Опрос состояния
                    switch (ImApplication.Status)
                    {
                        case ImBaseLoadStatus.IST_READY:
                            Done = true;
                            break; // Система готова
                        case ImBaseLoadStatus.IST_WAITFORLOGIN:
                        case ImBaseLoadStatus.IST_INTERNALLOADING:
                            // Система ожидает ввода пароля или загружается. Надо подождать.
                            System.Threading.Thread.Sleep(10);
                            break;
                    } // switch
                } // while
            } // try
            catch (Exception e)
            {
                Debug.WriteLine("Exception: " + e.Message);

                // System.Windows.Forms.MessageBox.(e);

                return false;
            }
            Debug.WriteLine(ImApplication.Status);
            return true;

        }

    }
}
