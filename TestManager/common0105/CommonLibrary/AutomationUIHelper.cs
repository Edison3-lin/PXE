/*
* AutomationUIHelper.cs 
* Use Automation library to find UI elements
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Automation;

namespace CaptainWin.CommonAPI
{
    /// <summary>
    ///  A helper help you to get the screen element's position, names. Can use class name or automation ID to Query.
    ///  need to use windows sdk tool inspect.exe to under stand the target element's class name or automation ID.
    /// </summary>
    public static class AutomationUIHelper
    {

        /// <summary>
        ///  Find element with element's class name and return the name
        /// </summary>
        /// <param name="programClassName">class name for root's children, and the target elements are under this node.</param>
        /// <param name="underProgramClassName">Target element's class name</param>
        /// <returns>List string of the name that the same class name find. </returns>
        public static List<string> GetNamesFromClassName(string programClassName, string underProgramClassName)
        {
            List<string> returnItems = new List<string>();

            var condition = new PropertyCondition(AutomationElement.ClassNameProperty, programClassName); //"Shell_TrayWnd"
            AutomationElement taskbarElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);

            if (taskbarElement != null)
            {
                condition = new PropertyCondition(AutomationElement.ClassNameProperty, underProgramClassName); //"Taskbar.TaskListButtonAutomationPeer"
                var taskListElements = taskbarElement.FindAll(TreeScope.Descendants, condition);
                string aa = "";
                foreach (AutomationElement aE in taskListElements)
                {
                    //     Console.WriteLine($"Element Name: {AE.Current.Name}");
                    returnItems.Add(aE.Current.Name);
                }
            }
            return returnItems;
        }

        /// <summary>
        ///  Use element's Clasnn name to find position
        /// </summary>
        /// <param name="programClassName">class name for root's children, and the target elements are under this node.</param>
        /// <param name="elementClassName">Target element's Class Name</param>
        /// <param name="FindIndex">use the class name may find multiple elements,
        /// default not passing this parameter will return the first one element's position.
        /// Or indicate which element by passing the index for this parameter.</param>
        /// <returns> Position for the finding element</returns>
        public static MousePosition GetAutomationElementPosition_ClassName(string programClassName, string elementClassName, int FindIndex = 0)
        {
            MousePosition rp = new MousePosition();
            AutomationElement rootChildElement = AutomationElement.RootElement.FindFirst(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.ClassNameProperty, programClassName));

            if (rootChildElement != null)
            {
                Condition classCondition = new PropertyCondition(AutomationElement.ClassNameProperty, elementClassName);
                //   Condition automationIdCondition = new PropertyCondition(AutomationElement.AutomationIdProperty, "");
                //   Condition combinedCondition = new AndCondition(classCondition, automationIdCondition);
                try
                {
                    AutomationElementCollection targetSet = rootChildElement.FindAll(TreeScope.Descendants, classCondition);
                    AutomationElement targetElement = targetSet[FindIndex];

                    if (targetElement != null)
                    {
                        return GetElementPosition(targetElement.Current.BoundingRectangle);
                    }
                }
                catch (Exception ex)
                {

                }



            }
            return rp;
        }
        /// <summary>
        ///  Use element's automation ID to find position
        /// </summary>
        /// <param name="programClassName">class name for root's children, and the target elements are under this node.</param>
        /// <param name="elementAutomationID">Target element's automation ID</param>
        /// <returns>Position for the finding element</returns>
        public static MousePosition GetAutomationElementPosition_AutomationID(string programClassName, string elementAutomationID)
        {
            MousePosition rp = new MousePosition();
            AutomationElement rootChildElement = AutomationElement.RootElement.FindFirst(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.ClassNameProperty, programClassName));

            if (rootChildElement != null)
            {
                Condition classCondition = new PropertyCondition(AutomationElement.AutomationIdProperty, elementAutomationID);
                //   Condition automationIdCondition = new PropertyCondition(AutomationElement.AutomationIdProperty, "");
                //   Condition combinedCondition = new AndCondition(classCondition, automationIdCondition);
                try
                {
                    AutomationElement targetElement = rootChildElement.FindFirst(TreeScope.Descendants, classCondition);

                    if (targetElement != null)
                    {
                        return GetElementPosition(targetElement.Current.BoundingRectangle);
                    }
                }
                catch (Exception ex)
                {

                }
            }
            return rp;
        }
        /// <summary>
        ///  List Automation UI root's children
        /// </summary>       
        /// <returns>List of string. Root children's class name</returns>
        public static List<string> ListRootChildrenClassName()
        {
            List<string> rootChildren = new List<string>();
            AutomationElement rootElement = AutomationElement.RootElement;
            AutomationElementCollection children = rootElement.FindAll(TreeScope.Children, Condition.TrueCondition);

            foreach (AutomationElement child in children)
            {
                //  Console.WriteLine($"Class Name for Child: {child.Current.ClassName}");
                rootChildren.Add(child.Current.ClassName);
            }
            return rootChildren;
        }


        private static MousePosition GetElementPosition(Rect ElementAera)
        {
            return new MousePosition
            {
                X = Convert.ToInt32(ElementAera.Left + ((ElementAera.Right - ElementAera.Left) / 2)),
                Y = Convert.ToInt32(ElementAera.Top + ((ElementAera.Bottom - ElementAera.Top) / 2))
            };
        }

    }
}
