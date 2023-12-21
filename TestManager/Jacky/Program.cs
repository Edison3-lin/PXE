using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Globalization;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;


namespace Jacky {
  public class Program {

    /// <summary>
    /// Sleep the system
    /// </summary>
    /// <param sleepParm="type count" > Sleep Type </param>
    public static void Sleep (string[] args) {
      string executablePath = @"c:\TestManager\ItemDownload\pwrtest.exe";
      string arguments = string.Format("/sleep /c:{1} /s:{0} /d:30 /p:40", args[0], args[1]);

      ProcessStartInfo startInfo = new ProcessStartInfo(executablePath)
      {
          Arguments = arguments,
          WorkingDirectory = @"c:\TestManager\ItemDownload",
          Verb = "runas"
      };

      try
      {
          Process process = new Process
          {
              StartInfo = startInfo
          };
          process.Start();
          process.WaitForExit();
      }
      catch (Exception ex)
      {
          Console.WriteLine("Error: " + ex.Message);
      }
    }  //Sleep

    /// <summary>
    /// Sleep the system
    /// </summary>
    /// <param sleepParm="type count" > Sleep Type </param>
    public static void Culture() {
        CultureInfo installedUICulture = CultureInfo.InstalledUICulture;
        Console.WriteLine("Installed UI Culture:");
        Console.WriteLine("Name: " + installedUICulture.Name);
        Console.WriteLine("DisplayName: " + installedUICulture.DisplayName);
        Console.WriteLine("EnglishName: " + installedUICulture.EnglishName);
        Console.WriteLine("TwoLetterISOLanguageName: " + installedUICulture.TwoLetterISOLanguageName);
        Console.WriteLine("ThreeLetterISOLanguageName: " + installedUICulture.ThreeLetterISOLanguageName);
    }

    /// <summary>
    /// Sleep the system
    /// </summary>
    /// <param sleepParm="type count" > Sleep Type </param>

/* WMI的常用类明细如下：
1、硬件类
冷却类别
Win32_Fan--风扇
Win32_HeatPipe--热管
Win32_Refrigeration--致冷
Win32_TemperatureProbe--温度传感
输入设备类别
Win32_Keyboard--键盘　
Win32_PointingDevice--指示设备（如鼠标）　
大容量存储类别
Win32_AutochkSetting--磁盘自动检查操作设置　
Win32_CDROMDrive--光盘驱动器　
Win32_DiskDrive--硬盘驱动器　
Win32_FloppyDrive--软盘驱动器　
Win32_PhysicalMedia--物理媒体　
Win32_TapeDrive--磁带驱动器　
主板、控制器、端口类别
Win32_1394Controller--1394控制器　
Win32_1394ControllerDevice--1394控制器设备
Win32_AllocatedResource--已分配的资源
Win32_AssociatedProcessorMemory--处理器和高速缓冲存储器
Win32_BaseBoard--主板
Win32_BIOS--BIOS（基本输入输出系统）
Win32_Bus--总线
Win32_CacheMemory--缓存内存
Win32_ControllerHasHub--USB控制器
Win32_DeviceBus--设备总线
Win32_DeviceMemoryAddress--设备存储器地址
Win32_DeviceSettings--设备设置
Win32_DMAChannel--DMA通道
Win32_FloppyController--软盘控制器
Win32_IDEController--IDE控制器
Win32_IDEControllerDevice--IDE控制器设备
Win32_InfraredDevice--红外线设备
Win32_IRQResource--中断（IRQ）资源
Win32_MemoryArray--内存数组
Win32_MemoryArrayLocation--内存数组位置
Win32_MemoryDevice--内存设备
Win32_MemoryDeviceArray--内存设备数组
Win32_MemoryDeviceLocation--内存设备位置
Win32_MotherboardDevice--主板设备
Win32_OnBoardDevice--插件设备
Win32_ParallelPort--并行端口
Win32_PCMCIAController--PCMCIA控制器
Win32_PhysicalMemory--物理内存
Win32_PhysicalMemoryArray--物理内存数组
Win32_PhysicalMemoryLocation--物理内存位置
Win32_PNPAllocatedResource--PNP保留资源
Win32_PNPDevice--PNP设备
Win32_PNPEntity--PNP实体
Win32_PortConnector--端口连接器
Win32_PortResource--端口资源
Win32_Processor--（CPU）处理器
Win32_SCSIController--SCSI控制器
Win32_SCSIControllerDevice--SCSI控制器设备
Win32_SerialPort--串行端口
Win32_SerialPortConfiguration--串行端口配置
Win32_SerialPortSetting--串行端口设置
Win32_SMBIOSMemory--内存有关的设备的管理
Win32_SoundDevice--声卡
Win32_SystemBIOS--系统BIOS
Win32_SystemDriverPNPEntity--系统驱动器PNP实体
Win32_SystemEnclosure--系统封闭
Win32_SystemMemoryResource--系统内存资源
Win32_SystemSlot--系统插槽
Win32_USBController--USB控制器
Win32_USBControllerDevice--USB控制器设备
Win32_USBHub--USB集线器


建网设备类别
Win32_NetworkAdapter--网络适配器
Win32_NetworkAdapterConfiguration--网络适配器配置
Win32_NetworkAdapterSetting--网络适配器设置


电源类别
Win32_AssociatedBattery--联合电池组
Win32_Battery--电池
Win32_CurrentProbe--当前传感
Win32_PortableBattery--便携式电池
Win32_PowerManagementEvent--电池事件管理
Win32_UninterruptiblePowerSupply--UPS电源
Win32_VoltageProbe--电压探测


打印类别
Win32_DriverForDevice--驱动器设备
Win32_Printer--打印机
Win32_PrinterConfiguration--打印机配置
Win32_PrinterController--打印机控制器
Win32_PrinterDriver--打印机驱动器
Win32_PrinterDriverDll--打印机驱动器DLL
Win32_PrinterSetting--打印机设置
Win32_PrintJob--打印工作
Win32_TCPIPPrinterPort--TCPIP打印机端口


电话类别
Win32_POTSModem--POTS调制解调器（Modem）
Win32_POTSModemToSerialPort--POTS调制解调器串行端口


视频监视器类别
Win32_DesktopMonitor--即插即用监视器
Win32_DisplayConfiguration--显示配置
Win32_DisplayControllerConfiguration--显示控制器配置
Win32_VideoConfiguration--视频配置
Win32_VideoController--视频控制器
Win32_VideoSettings--视频设置


2、操作系统类
COM类别
Win32_ClassicCOMApplicationClasses--
Win32_ClassicCOMClass--
Win32_ClassicCOMClassSettings--
Win32_ClientApplicationSetting--
Win32_COMApplication--COM应用
Win32_COMApplicationClasses--
Win32_COMApplicationSettings--
Win32_COMClass--
Win32_ComClassAutoEmulator--
Win32_ComClassEmulator--
Win32_ComponentCategory--
Win32_COMSetting--
Win32_DCOMApplication--DCOM应用
Win32_DCOMApplicationAccessAllowedSetting--
Win32_DCOMApplicationSetting--
Win32_ImplementedCategory--


桌面类别
Win32_Desktop--桌面
Win32_Environment--环境
Win32_TimeZone--时区
Win32_UserDesktop--使用者桌面


驱动程序类别
Win32_DriverVXD--
Win32_SystemDriver--系统驱动程序


文件系统类别
Win32_CIMLogicalDeviceCIMDataFile--
Win32_Directory--
Win32_DirectorySpecification--
Win32_DiskDriveToDiskPartition--
Win32_DiskPartition--磁盘逻辑分区
Win32_DiskQuota--NTFS磁盘分区定额
Win32_LogicalDisk--逻辑磁盘分区
Win32_LogicalDiskRootDirectory--
Win32_LogicalDiskToPartition--
Win32_MappedLogicalDisk--映射逻辑磁盘
Win32_OperatingSystemAutochkSetting--
Win32_QuotaSetting--
Win32_ShortcutFile--
Win32_SubDirectory--
Win32_SystemPartitions--
Win32_Volume--
Win32_VolumeQuota--
Win32_VolumeQuotaSetting--
Win32_VolumeUserQuota--


作业对象类别
Win32_CollectionStatistics--
Win32_LUID--
Win32_LUIDandAttributes--
Win32_NamedJobObject--
Win32_NamedJobObjectActgInfo--
Win32_NamedJobObjectLimit--
Win32_NamedJobObjectLimitSetting--
Win32_NamedJobObjectProcess--
Win32_NamedJobObjectSecLimit--
Win32_NamedJobObjectSecLimitSetting--
Win32_NamedJobObjectStatistics--
Win32_SIDandAttributes--
Win32_TokenGroups--
Win32_TokenPrivileges--


存储页面文件类别
Win32_LogicalMemoryConfiguration--逻辑内存配置
Win32_PageFile--页面文件
Win32_PageFileElementSetting--
Win32_PageFileSetting--页面文件设置
Win32_PageFileUsage--页面文件使用
Win32_SystemLogicalMemoryConfiguration--


多媒体视听类别
Win32_CodecFile--编解码器文件


建网类别
Win32_ActiveRoute--活动路由
Win32_IP4PersistedRouteTable--
Win32_IP4RouteTable--路由表
Win32_IP4RouteTableEvent--
Win32_NetworkClient--
Win32_NetworkConnection--
Win32_NetworkProtocol--网络协议
Win32_NTDomain--
Win32_PingStatus--
Win32_ProtocolBinding--协议绑定


操作系统事件类别
Win32_ComputerShutdownEvent--
Win32_ComputerSystemEvent--
Win32_DeviceChangeEvent--
Win32_ModuleLoadTrace--
Win32_ModuleTrace--
Win32_ProcessStartTrace--
Win32_ProcessStopTrace--
Win32_ProcessTrace--
Win32_SystemConfigurationChangeEvent--
Win32_SystemTrace--
Win32_ThreadStartTrace--
Win32_ThreadStopTrace--
Win32_ThreadTrace--
Win32_VolumeChangeEvent--


Win32_BootConfiguration--引导配置
Win32_ComputerSystem--计算机系统
Win32_ComputerSystemProcessor--计算机系统处理器
Win32_ComputerSystemProduct--计算机系统产品
Win32_DependentService--信任的服务
Win32_LoadOrderGroup--装载顺序组
Win32_LoadOrderGroupServiceDependencies--
Win32_LoadOrderGroupServiceMembers--
Win32_OperatingSystem--操作系统
Win32_OperatingSystemQFE--
Win32_OSRecoveryConfiguration--操作系统恢复配置
Win32_QuickFixEngineering--
Win32_StartupCommand--启动命令
Win32_SystemBootConfiguration--
Win32_SystemDesktop--
Win32_SystemDevices--
Win32_SystemLoadOrderGroups--
Win32_SystemNetworkConnections--
Win32_SystemOperatingSystem--
Win32_SystemProcesses--
Win32_SystemProgramGroups--Windows开始程序组
Win32_SystemResources--
Win32_SystemServices--系统服务
Win32_SystemSetting--
Win32_SystemSystemDriver--
Win32_SystemTimeZone--系统时区
Win32_SystemUsers--系统用户


进程类别
Win32_Process--进程
Win32_ProcessStartup--
Win32_Thread--线程


注册类别
Win32_Registry--注册表
调试程序作业类别
Win32_CurrentTime--当前时间
Win32_ScheduledJob--


安全类别
Win32_AccountSID--
Win32_ACE--
Win32_LogicalFileAccess--
Win32_LogicalFileAuditing--
Win32_LogicalFileGroup--
Win32_LogicalFileOwner--
Win32_LogicalFileSecuritySetting--
Win32_LogicalShareAccess--
Win32_LogicalShareAuditing--
Win32_LogicalShareSecuritySetting--
Win32_PrivilegesStatus--
Win32_SecurityDescriptor--
Win32_SecuritySetting--
Win32_SecuritySettingAccess--
Win32_SecuritySettingAuditing--
Win32_SecuritySettingGroup--
Win32_SecuritySettingOfLogicalFile--
Win32_SecuritySettingOfLogicalShare--
Win32_SecuritySettingOfObject--
Win32_SecuritySettingOwner--
Win32_SID--
Win32_Trustee--
服务类别
Win32_BaseService--基本服务
Win32_Service--服务


共享类别
Win32_DFSNode--
Win32_DFSNodeTarget--
Win32_DFSTarget--
Win32_ServerConnection--
Win32_ServerSession--
Win32_ConnectionShare--
Win32_PrinterShare--
Win32_SessionConnection--
Win32_SessionProcess--
Win32_ShareToDirectory--
Win32_Share--共享文件夹


开始菜单类别
Win32_LogicalProgramGroup--Windows开始逻辑程序组
Win32_LogicalProgramGroupDirectory--Windows开始逻辑程序组目录Win32_LogicalProgramGroupItem--Windows开始逻辑程序组项
Win32_LogicalProgramGroupItemDataFile--Windows开始逻辑程序组项数据文件
Win32_ProgramGroup--Windows程序组
Win32_ProgramGroupContents--Windows程序组内容
Win32_ProgramGroupOrItem--Windows程序组或项
存储类别
Win32_ShadowBy--
Win32_ShadowContext--
Win32_ShadowCopy--
Win32_ShadowDiffVolumeSupport--
Win32_ShadowFor--
Win32_ShadowOn--
Win32_ShadowProvider--
Win32_ShadowStorage--
Win32_ShadowVolumeSupport--
Win32_Volume--
Win32_VolumeUserQuota--


用户类别
Win32_Account--帐户
Win32_Group--组
Win32_GroupInDomain--域中的组
Win32_GroupUser--组用户
Win32_LogonSession--登录会话
Win32_LogonSessionMappedDisk--
Win32_NetworkLoginProfile--
Win32_SystemAccount--系统账户
Win32_UserAccount--使用账户
Win32_UserInDomain--域中的用户
Windows NT的事件日志类别
Win32_NTEventlogFile--事件日志文件
Win32_NTLogEvent--日志事件
Win32_NTLogEventComputer--日志事件计算机
Win32_NTLogEventLog--日志事件日志
Win32_NTLogEventUser--


Windows产品激活类别
Win32_ComputerSystemWindowsProductActivationSetting--
Win32_Proxy--代理
Win32_WindowsProductActivation--Windows产品激活


3、安装应用程序类
Win32_ActionCheck--
Win32_ApplicationCommandLine--
Win32_ApplicationService--
Win32_Binary--
Win32_BindImageAction--
Win32_CheckCheck--
Win32_ClassInfoAction--
Win32_CommandLineAccess--
Win32_Condition--
Win32_CreateFolderAction--
Win32_DuplicateFileAction--
Win32_EnvironmentSpecification--
Win32_ExtensionInfoAction--
Win32_FileSpecification--
Win32_FontInfoAction--
Win32_IniFileSpecification--
Win32_InstalledSoftwareElement--
Win32_LaunchCondition--
Win32_ManagedSystemElementResource--
Win32_MIMEInfoAction--
Win32_MoveFileAction--
Win32_MSIResource--
Win32_ODBCAttribute--
Win32_ODBCDataSourceAttribute--
Win32_ODBCDataSourceSpecification--
Win32_ODBCDriverAttribute--
Win32_ODBCDriverSoftwareElement--
Win32_ODBCDriverSpecification--
Win32_ODBCSourceAttribute--
Win32_ODBCTranslatorSpecification--
Win32_Patch--
Win32_PatchFile--
Win32_PatchPackage--
Win32_Product--
Win32_ProductCheck--
Win32_ProductResource--
Win32_ProductSoftwareFeatures--
Win32_ProgIDSpecification--
Win32_Property--
Win32_PublishComponentAction--
Win32_RegistryAction--
Win32_RemoveFileAction--
Win32_RemoveIniAction--
Win32_ReserveCost--
Win32_SelfRegModuleAction--
Win32_ServiceControl--
Win32_ServiceSpecification--
Win32_ServiceSpecificationService--
Win32_SettingCheck--
Win32_ShortcutAction--
Win32_ShortcutSAP--
Win32_SoftwareElement--
Win32_SoftwareElementAction--
Win32_SoftwareElementCheck--
Win32_SoftwareElementCondition--
Win32_SoftwareElementResource--
Win32_SoftwareFeature--
Win32_SoftwareFeatureAction--
Win32_SoftwareFeatureCheck--
Win32_SoftwareFeatureParent--
Win32_SoftwareFeatureSoftwareElements--
Win32_TypeLibraryAction--
4、WMI服务管理类
WMI配置类别
Win32_MethodParameterClass--方法参数类
WMI管理类别
Win32_WMISetting--WMI设置
Win32_WMIElementSetting--WMI单元设置


5、性能计数器类
格式化性能计数器类别
Win32_PerfFormattedData--
Win32_PerfFormattedData_ASP_ActiveServerPages--
Win32_PerfFormattedData_ContentFilter_IndexingServiceFilter--
Win32_PerfFormattedData_ContentIndex_IndexingService--
Win32_PerfFormattedData_InetInfo_InternetInformationServicesGlobal--
Win32_PerfFormattedData_ISAPISearch_HttpIndexingService--
Win32_PerfFormattedData_MSDTC_DistributedTransactionCoordinator--
Win32_PerfFormattedData_NTFSDRV_SMTPNTFSStoreDriver--
Win32_PerfFormattedData_PerfDisk_LogicalDisk--
Win32_PerfFormattedData_PerfDisk_PhysicalDisk--
Win32_PerfFormattedData_PerfNet_Browser--
Win32_PerfFormattedData_PerfNet_Redirector--
Win32_PerfFormattedData_PerfNet_Server--
Win32_PerfFormattedData_PerfNet_ServerWorkQueues--
Win32_PerfFormattedData_PerfOS_Cache--
Win32_PerfFormattedData_PerfOS_Memory--
Win32_PerfFormattedData_PerfOS_Objects--
Win32_PerfFormattedData_PerfOS_PagingFile--
Win32_PerfFormattedData_PerfOS_Processor--
Win32_PerfFormattedData_PerfOS_System--
Win32_PerfFormattedData_PerfProc_FullImage_Costly--
Win32_PerfFormattedData_PerfProc_Image_Costly--Win32_PerfFormattedData_PerfProc_JobObject--Win32_PerfFormattedData_PerfProc_JobObjectDetails--Win32_PerfFormattedData_PerfProc_Process--Win32_PerfFormattedData_PerfProc_ProcessAddressSpace_Costly--Win32_PerfFormattedData_PerfProc_Thread--Win32_PerfFormattedData_PerfProc_ThreadDetails_Costly--Win32_PerfFormattedData_PSched_PSchedFlow--
Win32_PerfFormattedData_PSched_PSchedPipe--Win32_PerfFormattedData_RemoteAccess_RASPort--Win32_PerfFormattedData_RemoteAccess_RASTotal--Win32_PerfFormattedData_RSVP_ACSRSVPInterfaces--Win32_PerfFormattedData_RSVP_ACSRSVPService--Win32_PerfFormattedData_SMTPSVC_SMTPServer--Win32_PerfFormattedData_Spooler_PrintQueue--
Win32_PerfFormattedData_TapiSrv_Telephony--
Win32_PerfFormattedData_Tcpip_ICMP--
Win32_PerfFormattedData_Tcpip_IP--
Win32_PerfFormattedData_Tcpip_NBTConnection--Win32_PerfFormattedData_Tcpip_NetworkInterface--
Win32_PerfFormattedData_Tcpip_TCP--
Win32_PerfFormattedData_Tcpip_UDP--Win32_PerfFormattedData_TermService_TerminalServices--Win32_PerfFormattedData_TermService_TerminalServicesSession--Win32_PerfFormattedData_W3SVC_WebService--


原始性能计数器类别
Win32_PerfRawData--Win32_PerfRawData_ASP_ActiveServerPages--Win32_PerfRawData_ContentFilter_IndexingServiceFilter--Win32_PerfRawData_ContentIndex_IndexingService--Win32_PerfRawData_InetInfo_InternetInformationServicesGlobal--Win32_PerfRawData_ISAPISearch_HttpIndexingService--Win32_PerfRawData_MSDTC_DistributedTransactionCoordinator--Win32_PerfRawData_NTFSDRV_SMTPNTFSStoreDriver--
Win32_PerfRawData_PerfDisk_LogicalDisk--
Win32_PerfRawData_PerfDisk_PhysicalDisk--
Win32_PerfRawData_PerfNet_Browser--
Win32_PerfRawData_PerfNet_Redirector--
Win32_PerfRawData_PerfNet_Server--
Win32_PerfRawData_PerfNet_ServerWorkQueues--
Win32_PerfRawData_PerfOS_Cache--
Win32_PerfRawData_PerfOS_Memory--
Win32_PerfRawData_PerfOS_Objects--
Win32_PerfRawData_PerfOS_PagingFile--
Win32_PerfRawData_PerfOS_Processor--
Win32_PerfRawData_PerfOS_System--
Win32_PerfRawData_PerfProc_FullImage_Costly--
Win32_PerfRawData_PerfProc_Image_Costly--
Win32_PerfRawData_PerfProc_JobObject--
Win32_PerfRawData_PerfProc_JobObjectDetails--
Win32_PerfRawData_PerfProc_Process--Win32_PerfRawData_PerfProc_ProcessAddressSpace_Costly--Win32_PerfRawData_PerfProc_Thread--
Win32_PerfRawData_PerfProc_ThreadDetails_Costly--
Win32_PerfRawData_PSched_PSchedFlow--
Win32_PerfRawData_PSched_PSchedPipe--
Win32_PerfRawData_RemoteAccess_RASPort--
Win32_PerfRawData_RemoteAccess_RASTotal--
Win32_PerfRawData_RSVP_ACSRSVPInterfaces--
Win32_PerfRawData_RSVP_ACSRSVPService--
Win32_PerfRawData_SMTPSVC_SMTPServer--
Win32_PerfRawData_Spooler_PrintQueue--
Win32_PerfRawData_TapiSrv_Telephony--
Win32_PerfRawData_Tcpip_ICMP--
Win32_PerfRawData_Tcpip_IP--
Win32_PerfRawData_Tcpip_NBTConnection--
Win32_PerfRawData_Tcpip_NetworkInterface--
Win32_PerfRawData_Tcpip_TCP--
Win32_PerfRawData_Tcpip_UDP--
Service_TerminalServices--Win32_PerfRawData_TermService_TerminalServicesSession--Win32_PerfRawData_W3SVC_WebService--

WMI - End */

    public static void Sysinfo(string className) {
        // string query = string.Format("SELECT * FROM {0}", wmi);
        // ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
        // ManagementObjectCollection queryCollection = searcher.Get();
        // foreach (ManagementObject m in queryCollection)
        // {
        //     Console.WriteLine("Operating System: " + m["Caption"].ToString());
        //     Console.WriteLine("Version: " + m["Version"].ToString());
        // }

        // string className = "Win32_OperatingSystem"; // 替换为您要查看的WMI类的名称

        ManagementClass mgmtClass = new ManagementClass(className);

        try
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT * FROM {className}");            
            ManagementObjectCollection collection = searcher.Get(); // 获取类的信息

            Console.WriteLine($"Class Name: {mgmtClass["__CLASS"]}");
            Console.WriteLine($"Description: {mgmtClass["__CLASS"]}");
            Console.WriteLine("Properties:");

            foreach (PropertyData prop in mgmtClass.Properties)
            {
                // Console.WriteLine($"  {prop.Name}: {prop.Type}");

                foreach (ManagementObject obj in collection)
                {
                    if (obj[prop.Name] != null)
                    {
                        string propertyValue = obj[prop.Name].ToString();
                        Console.WriteLine($"{prop.Name}: {propertyValue}");
                    }
                    else
                    {
                        Console.WriteLine($"{prop.Name}: Not available");
                    }
                }                
                // if(prop.Type == CimType.String)
                // {
                //     Console.WriteLine(prop.Value);
                //     Console.ReadKey();
                // }
            }

            Console.WriteLine("Methods:");

            foreach (MethodData method in mgmtClass.Methods)
            {
                Console.WriteLine($"  {method.Name}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void WMI_Run(string className, string propName, string propVal, string methodName) {
        try
        {
            ManagementClass mgmtClass = new ManagementClass(className);
            ManagementBaseObject inParams = mgmtClass.GetMethodParameters(methodName);

            // 设置方法的输入参数
            // "cmd.exe"：启动命令提示符。
            // "explorer.exe"：启动Windows资源管理器。
            // "calc.exe"：启动Windows计算器。
            // "mspaint.exe"：启动Windows画图工具。
            // "iexplore.exe"：启动Internet Explorer浏览器。
            inParams[propName] = propVal; // 根据方法的参数设置值

            // 调用方法
            ManagementBaseObject outParams = mgmtClass.InvokeMethod(methodName, inParams, null);

            // 检查方法执行结果
            if ((uint)outParams["ReturnValue"] == 0)
            {
                Console.WriteLine("Method executed successfully.");
            }
            else
            {
                Console.WriteLine("Method execution failed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
    static void Main(string[] args)
    {
        // string [] a = new string[]{args[0], args[1]};
        // Sleep(a);
        
        // Culture();

        // string className = args[0]; //"Win32_SystemDriver"; // 替换为您要查看的WMI类的名称

        // string className = "Win32_USBController"; // 替换为您要查看的WMI类的名称
        // Sysinfo(className);

        // string className2 = "Win32_Process"; // 替换为您要调用方法的WMI类的名称
        // string propName = "CommandLine"; // 替换为您要调用的方法名称
        // string methodName = "Create"; // 替换为您要调用的方法名称
        // string propVal = "calc.exe";
        // WMI_Run(className2, propName, propVal, methodName);

        // MySystem.get_Disk_VolumeSerialNumber();
        // Console.WriteLine(MySystem.get_OSVersion());
        // Console.WriteLine(a.cpu类别);
        // Console.WriteLine(a.cpu等级);
        // Console.WriteLine(a.cpu修正);

        Console.WriteLine(SystemInfo.GetPhysicalMemory());
    }

  } //Program
}
