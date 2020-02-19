using System;
using System.Diagnostics;
using E3Namespace;
using ImBaseExtensionDLL;

namespace AVSFromE3
{   
        class Program
        {            

            static readonly DateTime TimeStart = DateTime.Now;
            static void Main()
            {
                Debug.WriteLine(TimeStart);

          //  if (ImBaseEx.CheckImbaseConnection())
          //  {
                E3Devices.GetDevices();

                Debug.WriteLine(E3Devices.devList[1]);
                Debug.WriteLine(DateTime.Now - TimeStart);

               // ImBaseEx.CloseImbaseConnection();
          //  }
           
            /*
                if (ImBaseEx.CheckImbaseConnection())
                 {
                     foreach (string ImCode in ImCodes)
                     {
                         Debug.WriteLine(ImBaseEx.GetFullDesignation(ImCode));
                         Debug.WriteLine(DateTime.Now - TimeStart);
                     }
                     ImBaseEx.CloseImbaseConnection();               
                 }
                 
                */
        }
        }
    }


/*
               int Count = ImbApplication.Catalogs.Count;
               Debug.WriteLine(Count);

               Catalogs = ImbApplication.Catalogs;

               Debug.WriteLine(Catalogs.Count);
               ImbApplication.Restore();

               // Получить Каталог по имени
               Catalog = Catalogs.Item("Конструкторский");
               Debug.WriteLine(Catalog.Name);
               // Получить Каталог по имени таблицы
               Catalog = Catalogs.Item("Элтехника");
               Debug.WriteLine(Catalog.Name);
               // Получить Каталог по имени таблицы
               Folder = Catalog.FindFolder("Прочие изделия", ImFindObject.IFO_NAME);
               if (Folder != null) 
                   Debug.WriteLine(Folder.Name);
                  */
/*
//Индикатор ; УВНУ-6-35ДK; 900мм         i60901060803BB000008
ImDataBase.GetKeyData("i60901060803BB000008", out string TableRec);
Debug.WriteLine(TableRec);       

Debug.WriteLine(ImDataBase.GetKeyDataEx("i60901060803BB000008"));   
*/

/*
               GuidAttribute IMyInterfaceAttribute = (GuidAttribute)Attribute.GetCustomAttribute(typeof(IImbaseCatalogs), typeof(GuidAttribute));
               Debug.WriteLine("IMyInterface Attribute: " + IMyInterfaceAttribute.Value);
               Guid myGuid = new Guid(IMyInterfaceAttribute.Value);

               // Use the CLSID to instantiate the COM object using interop.
               Type type = Type.GetTypeFromCLSID(myGuid);
               Object comObj = Activator.CreateInstance(type);

               // Return a pointer to the objects IUnknown interface.
               IntPtr pIUnk = Marshal.GetIUnknownForObject(comObj);
               IntPtr pInterface;
               Int32 result = Marshal.QueryInterface(pIUnk, ref myGuid, out pInterface);



               int Count = iCatalogs.Count;
              for (int Index = 0; Index < Count; Index++)
              {
                  // Перебор всех Каталогов последовательно от первого до последнего
                  iCatalog = iCatalogs.Item(Index);
                  Debug.WriteLine(iCatalog.Name);
              } 


               */
