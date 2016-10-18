// <copyright file="ClipCursor.cs">
// Copyright (c) 2016 All Right Reserved
//
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
// </copyright>
// 
// <author>DeadlySurprise</author>
// <author>b2ccf637a919f68258cf3bfa7599fd80</author>
// <email>d3adlysurprise@gmail.com</email>
// <date>2016-10-18</date>
// <summary>Restricts the cursor from leaving the client area of a user selected window.</summary>

// Feel free to reuse or modify any of this code. 
// Write an email in case of questions or problems.

namespace ClipCursor
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;

    /// <summary>
    /// Contains methods for constraining the cursor to a window as well as a command-line tool for this task.
    /// </summary>
    public class ClipCursor_Main
    {    
        private const int WindowTitleMaxLength = 50; // Length window titles get truncated to
        private const int ValidateHandleThreshold = 10; // How often the user selected window handle gets validate
        private const int ClippingRefreshInterval = 100; // How often the clipped area is refreshed in milliseconds

        #region EnumerationsAndFlags
        private enum GetWindowLongIndex : int
        {
            GWL_WNDPROC = -4, GWL_HINSTANCE = -6, GWL_HWNDPARENT = -8, GWL_STYLE = -16, GWL_EXSTYLE = -20, GWL_USERDATA = -21, GWL_ID = -12
        }

        [Flags]
        private enum WindowStyles : int
        {
            WS_OVERLAPPED = 0x00000000, WS_POPUP = -2147483648, WS_CHILD = 0x40000000, WS_MINIMIZE = 0x20000000,
            WS_VISIBLE = 0x10000000, WS_DISABLED = 0x08000000, WS_CLIPSIBLINGS = 0x04000000, WS_CLIPCHILDREN = 0x02000000,
            WS_MAXIMIZE = 0x01000000, WS_CAPTION = 0x00C00000, WS_BORDER = 0x00800000, WS_DLGFRAME = 0x00400000,
            WS_VSCROLL = 0x00200000, WS_HSCROLL = 0x00100000, WS_SYSMENU = 0x00080000, WS_THICKFRAME = 0x00040000,
            WS_GROUP = 0x00020000, WS_TABSTOP = 0x00010000, WS_MINIMIZEBOX = 0x00020000, WS_MAXIMIZEBOX = 0x00010000
        }

        private enum SystemMetric : int
        {
            SM_CXBORDER = 5, SM_CYBORDER = 6, SM_CXSIZEFRAME = 32, SM_CYSIZEFRAME = 33, SM_CYCAPTION = 4, SM_CXFIXEDFRAME = 7, SM_CYFIXEDFRAME = 8
        }
        #endregion

        /// <summary>
        /// Runs the command line program.
        /// </summary>
        public static void Main()
        {
            List<IntPtr> windowHandles = null;
            int selectedIndex;
            string selectedIndexStr;
            IntPtr selectedWindowHandle;
            string selectedWindowTitle = string.Empty;

            while (true)
            {
                Console.WriteLine("Available windows : ");
                windowHandles = GetAllWindowHandles();
                Console.Write("Select a window by entering its index number : ");
                selectedIndexStr = Console.ReadLine();

                // Validate user choice
                if (!int.TryParse(selectedIndexStr, out selectedIndex) ||
                    selectedIndex < 1 ||
                    selectedIndex > windowHandles.Count)
                {
                    Console.Clear();
                    Console.WriteLine("Only use numbers that are on the list!");
                    continue;
                }

                selectedWindowHandle = windowHandles[selectedIndex - 1];
                selectedWindowTitle = GetWindowText(selectedWindowHandle, WindowTitleMaxLength);
                if (selectedWindowTitle == null)
                {
                    Console.WriteLine("The selected Window doesn't exists anymore!");
                    continue;
                }

                Console.WriteLine("Locking Cursor to \"{0:s}\"", selectedWindowTitle);

                LockCursor(selectedWindowHandle);
            }
        }

        /// <summary>
        /// Infinitely locks the cursor to the selected window if it has focus.
        /// Only returns once the window handle is destroyed.
        /// </summary>
        /// <param name="windowHandle">The handle of the window to which the cursor will be locked.</param>
        public static void LockCursor(IntPtr windowHandle)
        {
            Rectangle windowArea = new Rectangle();
            bool selectedWindowHadFocus = false;
            int validateHandleCount = 0;
            WindowStyles previousStyles = 0;
            Rectangle windowBorderSize = new Rectangle();
            string selectedWindowTitle = GetWindowText(windowHandle, WindowTitleMaxLength);

            while (true)
            {
                // Check if window styles changed so the program doesn't break if the window's borders style is changed
                if (previousStyles != GetWindowLong(windowHandle, GetWindowLongIndex.GWL_STYLE))
                {
                    // Determine border sizes for the selected window
                    windowBorderSize = GetWindowBorderSizes(windowHandle);
                    previousStyles = GetWindowLong(windowHandle, GetWindowLongIndex.GWL_STYLE);
                }

                if (GetForegroundWindow() == windowHandle)
                {
                    if (GetWindowRect(windowHandle, ref windowArea) == 0)
                    {
                        throw new Win32Exception(
                            Marshal.GetLastWin32Error(),
                            string.Format("Get window rectangle win32 error. selectedWindowHandle {0:d}", windowHandle));
                    }

                    windowArea.Left += windowBorderSize.Left;
                    windowArea.Top += windowBorderSize.Top;
                    windowArea.Bottom -= windowBorderSize.Bottom;
                    windowArea.Right -= windowBorderSize.Right;

                    if (ClipCursor(ref windowArea) == 0)
                    {
                        throw new Win32Exception(
                            Marshal.GetLastWin32Error(),
                            string.Format("Clip cursor win32 error. windowArea {0:s}", windowArea.ToString()));
                    }

                    selectedWindowHadFocus = true;
                    Thread.Sleep(500);
                }
                else if (selectedWindowHadFocus)
                {
                    // If the window lost focus remove the clipping area.
                    // Usually the clipping gets removed by default if the window loses focus. 
                    ClipCursor(IntPtr.Zero);
                    selectedWindowHadFocus = false;
                }

                // Validate the window every x amount of loops 
                validateHandleCount++;
                if (validateHandleCount > ValidateHandleThreshold)
                {
                    validateHandleCount = 0;
                    string tempWindowTitle = GetWindowText(windowHandle, WindowTitleMaxLength);
                    if (tempWindowTitle == null || tempWindowTitle != selectedWindowTitle)
                    {
                        Console.WriteLine("The selected Window doesn't exists anymore!");
                        ClipCursor(IntPtr.Zero);
                        break;
                    }
                }

                Thread.Sleep(ClippingRefreshInterval);
            }
        }

        /// <summary>
        /// Generate a list of all active window handles and optionally prints out their title texts in a numbered list.
        /// </summary>
        /// <param name="outputWindowNames">If true all window title texts are printed out.</param>
        /// <returns>Return a list all active window handles.</returns>
        public static List<IntPtr> GetAllWindowHandles(bool outputWindowNames = true)
        {
            Process[] processList;
            List<IntPtr> windowHandles = new List<IntPtr>();
            int indexCounter;

            // Print out (almost) every window title and save their handle
            processList = Process.GetProcesses();
            indexCounter = 1;

            if (windowHandles == null)
            {
                windowHandles = new List<IntPtr>();
            }
            else
            {
                windowHandles.Clear();
            }

            foreach (Process process in processList)
            {
                if (!string.IsNullOrEmpty(process.MainWindowTitle))
                {
                    if (outputWindowNames)
                    {
                        string windowTitle = RemoveSpecialCharacters(process.MainWindowTitle);
                        Console.WriteLine(
                            "({0:d}) : {1:s}",
                            indexCounter,
                            windowTitle.Substring(0, Math.Min(windowTitle.Length, WindowTitleMaxLength)));
                    }

                    windowHandles.Add(process.MainWindowHandle);
                    indexCounter++;
                }
            }

            return windowHandles;
        }

        /// <summary>
        /// Removes all escape and other non standard characters from the string so it can be safely printed to the console.
        /// </summary>
        /// <param name="str">The string to be sanitized.</param>
        /// <returns>Return the sanitized string.</returns>
        public static string RemoveSpecialCharacters(string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_. -]+", string.Empty, RegexOptions.Compiled);
        }

        /// <summary>
        /// Gets the size in pixel of a window's border.
        /// </summary>
        /// <param name="window">The handle of the window.</param>
        /// <returns>Returns the border size in pixel.</returns>
        public static Rectangle GetWindowBorderSizes(IntPtr window)
        {
            Rectangle windowBorderSizes = new Rectangle();

            WindowStyles styles = GetWindowLong(window, GetWindowLongIndex.GWL_STYLE);

            // Window has title-bar
            if (styles.HasFlag(WindowStyles.WS_CAPTION))
            {
                windowBorderSizes.Top += GetSystemMetrics(SystemMetric.SM_CYCAPTION);
            }

            // Window has re-sizable borders
            if (styles.HasFlag(WindowStyles.WS_THICKFRAME))
            {
                windowBorderSizes.Left += GetSystemMetrics(SystemMetric.SM_CXSIZEFRAME);
                windowBorderSizes.Right += GetSystemMetrics(SystemMetric.SM_CXSIZEFRAME);
                windowBorderSizes.Top += GetSystemMetrics(SystemMetric.SM_CYSIZEFRAME);
                windowBorderSizes.Bottom += GetSystemMetrics(SystemMetric.SM_CYSIZEFRAME);
            }
            else if (styles.HasFlag(WindowStyles.WS_BORDER) || styles.HasFlag(WindowStyles.WS_CAPTION))
            {
                // Window has normal borders
                windowBorderSizes.Left += GetSystemMetrics(SystemMetric.SM_CXFIXEDFRAME);
                windowBorderSizes.Right += GetSystemMetrics(SystemMetric.SM_CXFIXEDFRAME);
                windowBorderSizes.Top += GetSystemMetrics(SystemMetric.SM_CYFIXEDFRAME);
                windowBorderSizes.Bottom += GetSystemMetrics(SystemMetric.SM_CYFIXEDFRAME);
            }

            return windowBorderSizes;
        }

        /// <summary>
        /// Used to retrieve the title text of a window.
        /// </summary>
        /// <param name="hwnd">The handle of the window.</param>
        /// <param name="maxStringLength">The maximum length of the title string returned. Longer titles are truncated.</param>
        /// <returns>Return the title text of the window.</returns>
        private static string GetWindowText(IntPtr hwnd, int maxStringLength)
        {
            StringBuilder stringBuilder = new StringBuilder(maxStringLength);
            if (UnmanagedGetWindowText(hwnd, stringBuilder, maxStringLength) == 0)
            {
                return null;
            }

            return stringBuilder.ToString();
        }

        #region DLLImports
        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "GetWindowText")]
        private static extern int UnmanagedGetWindowText(IntPtr hwnd, StringBuilder lpString, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "GetForegroundWindow")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "GetWindowRect")]
        private static extern int GetWindowRect(IntPtr hwnd, ref Rectangle lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "GetClientRect")]
        private static extern int GetClientRect(IntPtr hwnd, ref Rectangle lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "ClipCursor")]
        private static extern int ClipCursor(ref Rectangle lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "ClipCursor")]
        private static extern int ClipCursor(IntPtr lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "GetSystemMetrics")]
        private static extern int GetSystemMetrics(SystemMetric index);

        [DllImport("user32.dll", CharSet = CharSet.Auto, EntryPoint = "GetWindowLong")]
        private static extern WindowStyles GetWindowLong(IntPtr hwnd, GetWindowLongIndex index);
        #endregion

        /// <summary>
        /// An implementation of the WINAPI RECT structure.
        /// </summary>
        public struct Rectangle
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            /// <summary>
            /// Generates a string containing all attributes of the rectangle.
            /// </summary>
            /// <returns>Returns a string containing all attributes of the rectangle.</returns>
            public override string ToString()
            {
                return string.Format("Left : {0:d}, Top : {1:d}, Right : {2:d}, Bottom : {3:d}", Left, Top, Right, Bottom);
            }
        }
    }
}
