using ImBase;
using System;
using System.Diagnostics;

namespace ImBaseExtensionDLL
{
    public static class ImBaseEx
    {
        public static ImbaseApplicationClass ImApplication = new ImbaseApplicationClass();
        public static ImDataBaseClass ImDataBase = new ImDataBaseClass();

        public static void CloseImbaseConnection()
        {
            ImApplication = null;
            ImDataBase = null;
        }
        //  public static ImbaseCatalogs Catalogs;       
        //  public static ImbaseCatalog Catalog;
        //  public static ImbaseFolder Folder;
        //  public static IImbaseKey ImBase;            

        public static string GetFullDesignation(string ImCode)
        {
            ImDataBase.GetKeyInfo(ImCode, out string TableRecord, out string CatalogRecord, out string KeysList);
            //      Debug.WriteLine(TableRecord);
            //      Debug.WriteLine(CatalogRecord);
            //      Debug.WriteLine(KeysList);

            int startIndex = CatalogRecord.IndexOf("\"ПОЛНОЕ ОБОЗНАЧЕНИЕ=") + 20;
            //  Debug.WriteLine(startIndex);

            int endIndex = CatalogRecord.IndexOf("\",КЛАСС=");
            //  Debug.WriteLine(endIndex);

            string FullDesignation = CatalogRecord.Substring(startIndex, endIndex - startIndex);

            return FullDesignation;
        }


        public static bool CheckImbaseConnection()
        {
            bool Done = false;
            try
            {
                // Создаем новое подключение, если указатель нулевой
                //
                // if (ImApplication == null) ImApplication = CoImbaseApplication.Create();
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
