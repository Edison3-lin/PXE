using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CaptainWin.CommonAPI{
    internal class SampleCodeAndDemo{
        public static string currentDirectory = Directory.GetCurrentDirectory() + '\\';
        public string GetLogPath(){
            //Reflection to common_update_logs.dll
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { };
                try{
                    Type typeGetTestLogPath = asmA.GetType("CaptainWin.CommonAPI.CommonUpdateLogs");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeGetTestLogPath);//create a instance
                    var miMethod = typeGetTestLogPath.GetMethod("GetTestLogPath");//miMethod = typeTest.GetMethod(dll_method)
                    var path = miMethod.Invoke(obj, p);
                    return path.ToString();
                }catch (Exception e){
                    throw e;
                }
            }
            return "";
        }
        public void ReadTestLogStatus(){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { };
                try{
                    Type typeReadTestLogStatus = asmA.GetType("CaptainWin.CommonAPI.CommonUpdateLogs");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeReadTestLogStatus);//create a instance
                    var miMethod = typeReadTestLogStatus.GetMethod("ReadTestLogStatus");//miMethod = typeTest.GetMethod(dll_method)
                    miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
        }
        public void WriteTestLogStatus(string writeTag, string p2){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { writeTag, p2 };
                try{
                    Type typeReadTestLogStatus = asmA.GetType("CaptainWin.CommonAPI.CommonUpdateLogs");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeReadTestLogStatus);//create a instance
                    var miMethod = typeReadTestLogStatus.GetMethod("WriteTestLogStatus");//miMethod = typeTest.GetMethod(dll_method)
                    miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
        }
        public void ReadJsonFile(string writeTag, string p2){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { writeTag, p2 };
                try{
                    Type typeReadTestLogStatus = asmA.GetType("CaptainWin.CommonAPI.CommonUpdateLogs");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeReadTestLogStatus);//create a instance
                    var miMethod = typeReadTestLogStatus.GetMethod("WriteTestLogStatus");//miMethod = typeTest.GetMethod(dll_method)
                    miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
        }

        public void DumpTestLogJsonFile(string filePath){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { filePath };
                try{
                    Type typeDumpTestLog = asmA.GetType("CaptainWin.CommonAPI.CommonUpdateLogs");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeDumpTestLog);//create a instance
                    var miMethod = typeDumpTestLog.GetMethod("DumpTestLogJsonFile");//miMethod = typeTest.GetMethod(dll_method)
                    miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
        }
        public string GetTimeZone(){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { };
                try{
                    Type typeGetTimeZone = asmA.GetType("CaptainWin.CommonAPI.CommonReadOOBESpecTable");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeGetTimeZone);//create a instance
                    var miMethod = typeGetTimeZone.GetMethod("GetTimeZone");//miMethod = typeTest.GetMethod(dll_method)
                    var timezone = miMethod.Invoke(obj, p);
                    return timezone.ToString();
                }catch (Exception e){
                    throw e;
                }
            }
            return "";
        }
        public void Setup(){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { };
                try{
                    Type typeSetup = asmA.GetType("CaptainWin.CommonAPI.CommonReadOOBESpecTable");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSetup);//create a instance
                    var miMethod = typeSetup.GetMethod("Setup");//miMethod = typeTest.GetMethod(dll_method)
                    miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
        }
        public string GetOOBESpec(string lang){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { lang };
                try{
                    Type typeGetOOBESpec = asmA.GetType("CaptainWin.CommonAPI.CommonReadOOBESpecTable");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeGetOOBESpec);//create a instance
                    var miMethod = typeGetOOBESpec.GetMethod("GetOOBESpec");//miMethod = typeTest.GetMethod(dll_method)
                    var timezone = miMethod.Invoke(obj, p);
                    return timezone.ToString();
                }
                catch (Exception e){
                    throw e;
                }
            }
            return "";
        }

        public bool CheckDeviceStatus(){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { };
                try{
                    Type typeDeviceStatus = asmA.GetType("CaptainWin.CommonAPI.CommonDevicesStatusCheck");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeDeviceStatus);//create a instance
                    var miMethod = typeDeviceStatus.GetMethod("CheckDeviceStatus");//miMethod = typeTest.GetMethod(dll_method)
                    var reslut = miMethod.Invoke(obj, p);
                    return (bool)reslut;
                }catch (Exception e){
                    throw e;
                }
            }
            return false;
        }
        public void SetupFullTable(string excelFile){
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { excelFile };
                try{
                    Type typeSetupFullTable = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSetupFullTable);//create a instance
                    var miMethod = typeSetupFullTable.GetMethod("SetupFullTable");//miMethod = typeTest.GetMethod(dll_method)
                    miMethod.Invoke(obj, p);
                }
                catch (Exception e){
                    throw e;
                }
            }
        }
        public List<string> GetSupportMetroAppList(string language){
            List<string> result = new List<string>();
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { language };
                try{
                    Type typeSupportMetroAppList = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSupportMetroAppList);//create a instance
                    var miMethod = typeSupportMetroAppList.GetMethod("GetSupportMetroAppList");//miMethod = typeTest.GetMethod(dll_method)
                    result = (List<string>)miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
            return result;
        }
        public List<string> GetSupportdesktopAppList(string language){
            List<string> result = new List<string>();
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { language };
                try{
                    Type typeSupportMetroAppList = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSupportMetroAppList);//create a instance
                    var miMethod = typeSupportMetroAppList.GetMethod("GetSupportdesktopAppList");//miMethod = typeTest.GetMethod(dll_method)
                    result = (List<string>)miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
            return result;
        }
        public List<string> GetSupportdesktopBarAppList(string language){
            List<string> result = new List<string>();
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { language };
                try{
                    Type typeSupportMetroAppList = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSupportMetroAppList);//create a instance
                    var miMethod = typeSupportMetroAppList.GetMethod("GetSupportdesktopBarAppList");//miMethod = typeTest.GetMethod(dll_method)
                    result = (List<string>)miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
            return result;
        }
        public List<string> GetSupportShortCutAppList(string language){
            List<string> result = new List<string>();
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { language };
                try{
                    Type typeSupportMetroAppList = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSupportMetroAppList);//create a instance
                    var miMethod = typeSupportMetroAppList.GetMethod("GetSupportShortCutAppList");//miMethod = typeTest.GetMethod(dll_method)
                    result = (List<string>)miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
            return result;
        }
        public List<string> GetSupportBrowserAppList(string language){
            List<string> result = new List<string>();
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { language };
                try{
                    Type typeSupportMetroAppList = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSupportMetroAppList);//create a instance
                    var miMethod = typeSupportMetroAppList.GetMethod("GetSupportBrowserAppList");//miMethod = typeTest.GetMethod(dll_method)
                    result = (List<string>)miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
            return result;
        }
        public List<string> GetSupportOOBEIntegrationList(string language){
            List<string> result = new List<string>();
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { language };
                try{
                    Type typeSupportMetroAppList = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSupportMetroAppList);//create a instance
                    var miMethod = typeSupportMetroAppList.GetMethod("GetSupportBrowserAppList");//miMethod = typeTest.GetMethod(dll_method)
                    result = (List<string>)miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
            return result;
        }
        public List<string> GetSupportWindowsNextRecommendedList(string language){
            List<string> result = new List<string>();
            Assembly asmA = Assembly.LoadFrom(currentDirectory + '\\' + "CommonLibrary.dll");

            if (null != asmA){

                object[] p = new object[] { language };
                try{
                    Type typeSupportMetroAppList = asmA.GetType("CaptainWin.CommonAPI.CommonReadGSAMRD");//typeTest = asmA.GetType(dll_namespace.dll_class)
                    object obj = Activator.CreateInstance(typeSupportMetroAppList);//create a instance
                    var miMethod = typeSupportMetroAppList.GetMethod("GetSupportWindowsNextRecommendedList");//miMethod = typeTest.GetMethod(dll_method)
                    result = (List<string>)miMethod.Invoke(obj, p);
                }catch (Exception e){
                    throw e;
                }
            }
            return result;
        }
        public void Run(){

            Console.WriteLine(currentDirectory);
            SampleCodeAndDemo generalConsoleApp = new SampleCodeAndDemo();
            string logpath = generalConsoleApp.GetLogPath();
            Console.WriteLine(logpath);
            generalConsoleApp.ReadTestLogStatus();
            generalConsoleApp.WriteTestLogStatus("TestStatus", "Done");
            generalConsoleApp.WriteTestLogStatus("TestResult", "PASS");
            generalConsoleApp.ReadTestLogStatus();
            generalConsoleApp.DumpTestLogJsonFile(logpath);

            //ReadOOBESPECbyCS readOOBESPECbyCS = new ReadOOBESPECbyCS();

            string timeZone = generalConsoleApp.GetTimeZone();

            generalConsoleApp.Setup();

            // Get the default system language
            CultureInfo systemCulture = CultureInfo.InstalledUICulture;

            // Display language information
            Console.WriteLine($"Default System Language: {systemCulture.DisplayName}");
            Console.WriteLine($"Language Code: {systemCulture.Name}");
            Console.WriteLine(systemCulture.ToString().ToUpper());
            string sample = systemCulture.ToString().ToUpper();
            char charToRemove = '-'; // Replace this with the character you want to remove

            // Remove the specified character using LINQ
            string modifiedString = new string(sample.Where(c => c != charToRemove).ToArray());
            Console.WriteLine(modifiedString);

            string specTimeZone = generalConsoleApp.GetOOBESpec(modifiedString);

            if (specTimeZone.IndexOf(timeZone) >= 0){
                Console.WriteLine("PASS, match Time Zone between OOBE SPEC and Win11");
            }else{
                Console.WriteLine("FAIL, not match Time Zone between OOBE SPEC and Win11");
            }

            generalConsoleApp.CheckDeviceStatus();
            generalConsoleApp.SetupFullTable("TEST_SCD_RV07RC.xls");
            generalConsoleApp.GetSupportMetroAppList("GB");
            generalConsoleApp.GetSupportdesktopAppList("GB");
            generalConsoleApp.GetSupportdesktopBarAppList("GB");
            generalConsoleApp.GetSupportBrowserAppList("GB");
            generalConsoleApp.GetSupportShortCutAppList("GB");
            generalConsoleApp.GetSupportOOBEIntegrationList("GB");
            generalConsoleApp.GetSupportWindowsNextRecommendedList("GB");

        }
    }
}