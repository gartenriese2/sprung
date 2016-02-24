using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows;
using System.Windows.Automation;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace Sprung
{

    //class WindowsByClassFinder {
    //    public delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lparam);

    //    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    //    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    //    [DllImport("user32.dll")]
    //    [return: MarshalAs(UnmanagedType.Bool)]
    //    public extern static bool EnumWindows(EnumWindowsDelegate lpEnumFunc, IntPtr lparam);

    //    [DllImport("User32", CharSet = CharSet.Auto, SetLastError = true)]
    //    public static extern int GetWindowText(IntPtr windowHandle, StringBuilder stringBuilder, int nMaxCount);

    //    [DllImport("user32.dll", EntryPoint = "GetWindowTextLength", SetLastError = true)]
    //    internal static extern int GetWindowTextLength(IntPtr hwnd);


    //    /// <summary>Find the windows matching the specified class name.</summary>

    //    public static IEnumerable<IntPtr> WindowsMatching(string className) {
    //        return new WindowsByClassFinder(className)._result;
    //    }

    //    private WindowsByClassFinder(string className) {
    //        _className = className;
    //        EnumWindows(callback, IntPtr.Zero);
    //    }

    //    private bool callback(IntPtr hWnd, IntPtr lparam) {
    //        if (GetClassName(hWnd, _apiResult, _apiResult.Capacity) != 0) {
    //            if (string.CompareOrdinal(_apiResult.ToString(), _className) == 0) {
    //                _result.Add(hWnd);
    //            }
    //        }

    //        return true; // Keep enumerating.
    //    }

    //    public static IEnumerable<string> WindowTitlesForClass(string className) {
    //        foreach (var windowHandle in WindowsMatchingClassName(className)) {
    //            int length = GetWindowTextLength(windowHandle);
    //            StringBuilder sb = new StringBuilder(length + 1);
    //            GetWindowText(windowHandle, sb, sb.Capacity);
    //            yield return sb.ToString();
    //        }
    //    }

    //    public static IEnumerable<IntPtr> WindowsMatchingClassName(string className) {
    //        if (string.IsNullOrWhiteSpace(className))
    //            throw new ArgumentOutOfRangeException("className", className, "className can't be null or blank.");

    //        return WindowsMatching(className);
    //    }

    //    private readonly string _className;
    //    private readonly List<IntPtr> _result = new List<IntPtr>();
    //    private readonly StringBuilder _apiResult = new StringBuilder(1024);
    //}

    class WindowManager
    {

        private Settings settings;
        private List<Window> windows = new List<Window>();
        private bool showTabs = false;
        private bool chromeDone = false;

        public WindowManager(Settings settings)
        {
            this.settings = settings;
        }

        public List<Window> getProcesses()
        {
            windows.Clear();
            EnumDelegate callback = new EnumDelegate(EnumWindowsProc);
            bool enumDel = EnumDesktopWindows(IntPtr.Zero, callback, IntPtr.Zero);
            if (!enumDel)
            {
                throw new Exception("Calling EnumDesktopWindows: Error ocurred: " + Marshal.GetLastWin32Error());
            }
            return windows;
        }

        public List<Window> getProcesses(bool showTabs)
        {
            this.showTabs = showTabs;
            List<Window> windows = getProcesses();
            this.showTabs = false;
            chromeDone = false;
            return windows;
        }

        //public IEnumerable<string> ChromeWindowTitles() {
        //    foreach (var title in WindowsByClassFinder.WindowTitlesForClass("Chrome_WidgetWin_0"))
        //        if (!string.IsNullOrWhiteSpace(title))
        //            yield return title;

        //    foreach (var title in WindowsByClassFinder.WindowTitlesForClass("Chrome_WidgetWin_1"))
        //        if (!string.IsNullOrWhiteSpace(title))
        //            yield return title;
        //}


        //[DllImport("oleacc.dll")]
        //public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, ref Guid refID, ref Accessibility.IAccessible ppvObject);

        //[DllImport("oleacc.dll")]
        //public static extern int AccessibleChildren(Accessibility.IAccessible paccContainer, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);

        //public static Accessibility.IAccessible GetObjectByName(Accessibility.IAccessible objParent,
        //      Regex objName, bool ignoreInvisible, int depth) {
        //    Accessibility.IAccessible objToReturn = default(Accessibility.IAccessible);
        //    if (objParent != null) {
        //        int _ChildCount;
        //        try {
        //            _ChildCount = objParent.accChildCount;
        //        } catch (Exception) {
        //            return objToReturn;
        //        }
        //        Accessibility.IAccessible[] children = new Accessibility.IAccessible[_ChildCount];
        //        int _out;
        //        AccessibleChildren(objParent, 0, _ChildCount - 1, children, out _out);
        //        foreach (Accessibility.IAccessible child in children) {
        //            string childName = null;
        //            string childState = string.Empty;

        //            try {
        //                childName = child.get_accName(0);
        //                //childState =
        //                //  GetStateText(Convert.ToUInt32(child.get_accState(0)));
        //            } catch (Exception) {
        //            }

        //            string output = "";
        //            for (int i = 0; i < depth; ++i) {
        //                output += "\t";
        //            }
        //            Debug.WriteLine(output + (childName == null ? "null" : childName));

        //            if (ignoreInvisible) {
        //                if (childName != null
        //                    && objName.Match(childName).Success
        //                    && !childState.Contains("invisible")) {
        //                    return child;
        //                }
        //            } else {
        //                if (childName != null
        //                    && objName.Match(childName).Success) {
        //                    return child;
        //                }
        //            }

        //            if (ignoreInvisible) {
        //                if (!childState.Contains("invisible")) {
        //                    objToReturn = GetObjectByName(child, objName, ignoreInvisible, depth + 1);
        //                    if (objToReturn != default(Accessibility.IAccessible)) {
        //                        return objToReturn;
        //                    }
        //                }
        //            } else {
        //                objToReturn = GetObjectByName(child, objName, ignoreInvisible, depth + 1);
        //                if (objToReturn != default(Accessibility.IAccessible)) {
        //                    return objToReturn;
        //                }
        //            }

        //        }
        //    }
        //    return objToReturn;
        //}

        private bool EnumWindowsProc(IntPtr hWnd, int lParam)
        {
            if (IsWindowVisible(hWnd)) {
                Window window = new Window(hWnd);
                if (!settings.isWindowTitleExcluded(window.getTitle()) && !window.hasNoTitle())
                {
                    if ((showTabs || settings.isListTabsAsWindows()) && window.getProcessName() == "firefox")
                    {
                        windows.AddRange(getFirefoxTabs(window));
                    }
                    else if ((showTabs || settings.isListTabsAsWindows()) && window.getProcessName() == "chrome" && !chromeDone) {
                        windows.AddRange(getChromeTabs(window));
                        chromeDone = true;

                        // TEST
                        //Debug.WriteLine("TEST TEST TEST!");
                        //var titles = ChromeWindowTitles();
                        //foreach (var title in titles) {
                        //    Debug.WriteLine(title);
                            
                        //}
                        //Guid iid = typeof(Accessibility.IAccessible).GUID;
                        //Accessibility.IAccessible ia = null;
                        //AccessibleObjectFromWindow(hWnd, 0, ref iid, ref ia);
                        //Regex regexName = new Regex("New Tab - Google Chrome");
                        //var acc = GetObjectByName(ia, regexName, false, 0);
                        //Debug.WriteLine(acc.accName);
      

                        
                        //Debug.WriteLine("ENDE ENDE ENDE!");

                    }
                    else if ((showTabs || settings.isListTabsAsWindows()) && window.getProcessName() == "iexplore")
                    {
                        windows.AddRange(getIETabs(window));
                    }
                    else
                    {
                        windows.Add(window);
                    }
                    
                }
            }
            return true;
        }

        private List<Window> getFirefoxTabs(Window firefoxWindow)
        {
            List<Window> tabs = new List<Window>();
            string path = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName;
            if (Environment.OSVersion.Version.Major >= 6) path = Directory.GetParent(path).ToString();
            path += "/AppData/Roaming/Mozilla/Firefox/Profiles";
            path = Directory.GetDirectories(path)[0];
            path += "/sessionstore-backups/";

            String recovery = path + "/recovery.js";
            String recoveryTmp = path + "/recovery.js.tmp";
            String file = File.Exists(recoveryTmp) ? recoveryTmp : recovery;

            StreamReader streamReader = new StreamReader(file);
            String content = streamReader.ReadToEnd();
            JObject data = JObject.Parse(content);
            int currentTabIndex = 0, i = 0;
            foreach (JObject tab in data["windows"][0]["tabs"])
            {
                String title = (String) tab["entries"].Last["title"];
                title += " - Mozilla Firefox";
                currentTabIndex = title == firefoxWindow.getTitle() ? i : currentTabIndex;
                tabs.Add(new FirefoxTabWindow(firefoxWindow.getHandle(), currentTabIndex, i++, title));
            }
            foreach(FirefoxTabWindow w in tabs) {
                w.currentTabIndex = currentTabIndex;
            }
            return tabs;
        }

        private List<Window> getChromeTabs(Window chromeWindow) {
            List<Window> tabs = new List<Window>();
            string path = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName;
            if (Environment.OSVersion.Version.Major >= 6) path = Directory.GetParent(path).ToString();

            var proc = chromeWindow.getProcess();
            if (proc.MainWindowHandle == IntPtr.Zero) {
                // Chrome process does not have a window
                return tabs;
            }

            var currentTabTitle = proc.MainWindowTitle;
            var automationElements = AutomationElement.FromHandle(proc.MainWindowHandle);

            // Find `New Tab` element
            var propCondNewTab = new PropertyCondition(AutomationElement.NameProperty, "New Tab");
            var elemNewTab = automationElements.FindFirst(TreeScope.Descendants, propCondNewTab);

            // Get parent of `New Tab` element
            var treeWalker = TreeWalker.ControlViewWalker;
            var elemTabStrip = treeWalker.GetParent(elemNewTab);

            // Loop through all tabs and sort them by x
            var tabItemCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
            SortedDictionary<double, AutomationElement> orderedTabItems = new SortedDictionary<double, AutomationElement>();
            foreach (AutomationElement tabItem in elemTabStrip.FindAll(TreeScope.Children, tabItemCondition)) {
                System.Windows.Rect rectangleProperty = (Rect)tabItem.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
                orderedTabItems.Add(rectangleProperty.X, tabItem);
            }

            Debug.WriteLine("Found " + orderedTabItems.Keys.Count + " chrome tabs");

            // calculate current index
            int currentTabIndex = 0;
            for (int i = 0; i < orderedTabItems.Keys.Count; i++) {
                var key = orderedTabItems.Keys.ElementAt(i);
                var tabItem = orderedTabItems[key];
                var nameProperty = tabItem.GetCurrentPropertyValue(AutomationElement.NameProperty);
                string title = nameProperty.ToString() + " - Google Chrome";
                currentTabIndex = title == currentTabTitle ? i : currentTabIndex;
                tabs.Add(new FirefoxTabWindow(chromeWindow.getHandle(), currentTabIndex, i, title));
            }

            // set current index to each tab
            foreach (FirefoxTabWindow w in tabs) {
                w.currentTabIndex = currentTabIndex;
            }

            return tabs;
        }

        public List<Window> getIETabs(Window ieWindow) {
            List<Window> tabs = new List<Window>();
            foreach (SHDocVw.InternetExplorer tab in new SHDocVw.ShellWindows())
            {
                if (!tab.LocationURL.StartsWith("file"))
                {
                    Console.WriteLine(tab.LocationURL);
                    Console.WriteLine("visible = " + tab.Visible);
                    Console.WriteLine("hasFocus = " + ((mshtml.HTMLDocument)tab.Document).hasFocus());
                    tabs.Add(new IETabWindow(tab));
                }
            }
            return tabs;
        }

        private delegate bool EnumDelegate(IntPtr hWnd, int lParam);

        [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindowVisible(IntPtr hWnd);

    }
}
