/*
* BrowserUIA.cs - Browser UI Automation Control funcitons
* CheckRootViewName - Check if root view name of a web page is correct
* CheckControlName - Check if the name of a control is in web page
* InvokeHyperlink - Invoke a hyperlink
* InvokeButton - Invoke a button
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Chris Huang   <Chris_Huang@quantatw.com>
*/

using System.Windows.Automation;


namespace CaptainWin.CommonAPI{
    /// <summary>
    /// Thie class utilizes UI Automation to check web pages name and invoke hyperlink
    /// </summary>
    /// 
    public class BrowserUIA {
        private static AutomationElement getBrowserRootElement(string name) {
            
            AutomationElement desktop = AutomationElement.RootElement;
            AutomationElement widgetWin;
            AutomationElement rootView = null;

            widgetWin = desktop.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty, "Chrome_WidgetWin_1"));
            if ( widgetWin.Current.Name.Contains(name) ) {
                rootView = widgetWin.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty, "BrowserRootView"));

            }
            return rootView;
        }
        /// <summary>
        /// Check if the name fo browser root view matches the webpage name we want to find.
        /// </summary>
        /// <param name="name">The name of a browser root view</param>
        /// <returns>true if browser root view with parameter name is found, false if not found</returns>
        public static bool CheckRootViewName(string name) {
            
            bool result = false;

            AutomationElement topView = getBrowserRootElement(name);

            if (topView != null) { 
                result = true;
            }
            return result;
        }
        /// <summary>
        /// Check if a control name exists in a browser root view.
        /// </summary>
        /// <param name="browserRootViewName">The name of a browser root view</param>
        /// <param name="name">The name of a control view</param>
        /// <returns>true if a control with parameter name is found, false if not</returns>
        public static int CheckControlName(string browserRootViewName, string name) {

            AutomationElement topView = getBrowserRootElement(browserRootViewName);

            AutomationElementCollection decendantViews = topView.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));

            return decendantViews.Count;
        }
        /// <summary>
        /// Invoke a Hyperlink with parameter name in a browser root view.
        /// </summary>
        /// <param name="browserRootViewName">The name of a browser root view</param>
        /// <param name="name">The name of the hyperlink </param>
        /// <returns>true if hyperlink is invoked, false if not</returns>
        public static bool InvokeHyperlink(string browserRootViewName, string name) {
            
            bool result = false;

            AutomationElement topView = getBrowserRootElement(browserRootViewName);
            AutomationElement decendantView;

            Condition condition = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, name), new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink));
            decendantView = topView.FindFirst(TreeScope.Descendants, condition);

            if (decendantView != null) {
                InvokePattern hyperlinkInvoke = (InvokePattern)decendantView.GetCurrentPattern(InvokePattern.Pattern);
                hyperlinkInvoke.Invoke();
                result = true;
            }
            return result;
        }
        /// <summary>
        /// Invoke a button with parameter name in a browser root view.
        /// </summary>
        /// <param name="browserRootViewName">The name of a browser root view</param>
        /// <param name="name">The name of a button view</param>
        /// <returns>true if the button is invoked, false if not</returns>
        public static bool InvokeButton(string browserRootViewName, string name) {
            
            bool result = false;

            AutomationElement topView = getBrowserRootElement(browserRootViewName);
            AutomationElement decendantView;

            Condition condition = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, name), new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            decendantView = topView.FindFirst(TreeScope.Descendants, condition);

            if (decendantView != null) {
                InvokePattern hyperlinkInvoke = (InvokePattern)decendantView.GetCurrentPattern(InvokePattern.Pattern);
                hyperlinkInvoke.Invoke();
                result = true;
            }
            return result;
        }
    }
}