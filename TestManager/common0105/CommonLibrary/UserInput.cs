/*
* Common.cs - Common API for test items
* UserInput.cs - Library for simulating user input
* Mouse_Functions - Perform mouse operations
* KB_Functions - Perform keyboard operations
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Jacky Kao   <Jacky.Kao@quantatw.com>
*/

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;

namespace CaptainWin.CommonAPI
{
    /// <summary>
    /// Thie class simulates user input
    /// </summary>
    public class UserInput {
    [DllImport("user32.dll")]
    private static extern void SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(int dwFlags, int dx, int dy, 
                                           int dwData, int dwExtraInfo);
    
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, 
                                           uint dwFlags, UIntPtr dwExtraInfo);
    
    [DllImport("user32.dll")]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hwnd, ref Rectangle rectangle);

    private const int _mouseLeftdown = 0x02;
    private const int _mouseLeftup = 0x04;
    private const int _mouseRightdown = 0x0008;
    private const int _mouseRightup = 0x0010;
    private const int _kbEnter = 0x0D;
    private const int _kbSpace = 0x20;
    private const int _kbLShift = 0xA0;
    private const int _kbRShift = 0xA1;
    private const int _kbLCTRL = 0xA2;
    private const int _kbRCTRL = 0xA3;
    private const int _kbLALT = 0xA4;
    private const int _kbRALT = 0xA5;
    private const int _kbLeft = 0x25;
    private const int _kbUp = 0x26;
    private const int _kbRight = 0x27;
    private const int _kbDown = 0x28;
    private const int _kbDelete = 0x2E;
    private const int _kbTab = 0x09;
    private const int _kbCapsLock = 0x14;
    private const int _kbEscape = 0x1B;
    private const int _kbBackspace = 0x08;
    private const int _kbUnicode = 0x0004;
    private const int _kbWin = 0x5B;
        
    /// <summary>
    /// Move mouse to x,y
    /// </summary>
    /// <param name="x">x coordinate of the cursor</param>
    /// <param name="y">y coordinate of the cursor</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_Move (int x, int y) {
      try {
        Thread.Sleep(50);
        SetCursorPos(x, y);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Click left button of mouse
    /// </summary>
    /// <param name="time">The time to hold the button</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_LeftClick(int time) {
      try {
        Thread.Sleep(50);
        mouse_event( _mouseLeftdown, 0, 0, 0, 0);
        Thread.Sleep(time);
        mouse_event( _mouseLeftup, 0, 0, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press and hold left button of mouse
    /// </summary>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_HoldLeft() {
      try {
        Thread.Sleep(50);
        mouse_event( _mouseLeftdown, 0, 0, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Release left button of mouse
    /// </summary>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_ReleaseLeft() {
      try {
        Thread.Sleep(50);
        mouse_event( _mouseLeftup, 0, 0, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }
    
    /// <summary>
    /// Click right button of mouse
    /// </summary>
    /// <param name="time">The time to hold the button</param>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_RightClick(int time) {
      try {
        Thread.Sleep(50);
        mouse_event( _mouseRightdown, 0, 0, 0, 0);
        Thread.Sleep(time);
        mouse_event( _mouseRightup, 0, 0, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press and hold right button of mouse
    /// </summary>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_HoldRight() {
      try {
        Thread.Sleep(50);
        mouse_event( _mouseRightdown, 0, 0, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Release right button of mouse
    /// </summary>
    /// <returns>0 if success, -1 if error</returns>
    public int Mouse_ReleaseRight() {
      try {
        Thread.Sleep(50);
        mouse_event( _mouseRightup, 0, 0, 0, 0);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Type a char simulating keyboard input
    /// </summary>
    /// <param name="key">The char to be pressed</param>
    /// <returns></returns>
    public int KB_TypeChar(char key) {
      try {
        Thread.Sleep(50);
        switch (key) {
          case ',':
            keybd_event(0, (byte)0xBC, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xBC, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '.':
            keybd_event(0, (byte)0xBE, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xBE, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '/':
            keybd_event(0, (byte)0xBF, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xBF, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '\\':
            keybd_event(0, (byte)0xDC, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xDC, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case ';':
            keybd_event(0, (byte)0xBA, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xBA, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '\'':
            keybd_event(0, (byte)0xDE, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xDE, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '[':
            keybd_event(0, (byte)0xDB, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xDB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case ']':
            keybd_event(0, (byte)0xDD, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xDD, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '{':
            keybd_event(0, (byte)0xDB, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xDB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '}':
            keybd_event(0, (byte)0xDD, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xDD, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '-':
            keybd_event(0, (byte)0xBD, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xBD, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '=':
            keybd_event(0, (byte)0xBB, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xBB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '+':
            keybd_event(0, (byte)0xBB, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)0xBB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          default:
            keybd_event(0, (byte)key, _kbUnicode, UIntPtr.Zero);
            keybd_event(0, (byte)key, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
        }
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press and hold a char simulating keyboard input
    /// </summary>
    /// <param name="key">The char to be pressed</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_HoldChar(char key) {
      try {
        Thread.Sleep(50);
        switch (key) {
          case ',':
            keybd_event(0, (byte)0xBC, _kbUnicode, UIntPtr.Zero);
            break;
          case '.':
            keybd_event(0, (byte)0xBE, _kbUnicode, UIntPtr.Zero);
            break;
          case '/':
            keybd_event(0, (byte)0xBF, _kbUnicode, UIntPtr.Zero);
            break;
          case '\\':
            keybd_event(0, (byte)0xDC, _kbUnicode, UIntPtr.Zero);
            break;
          case ';':
            keybd_event(0, (byte)0xBA, _kbUnicode, UIntPtr.Zero);
            break;
          case '\'':
            keybd_event(0, (byte)0xDE, _kbUnicode, UIntPtr.Zero);
            break;
          case '[':
            keybd_event(0, (byte)0xDB, _kbUnicode, UIntPtr.Zero);
            break;
          case ']':
            keybd_event(0, (byte)0xDD, _kbUnicode, UIntPtr.Zero);
            break;
          case '{':
            keybd_event(0, (byte)0xDB, _kbUnicode, UIntPtr.Zero);
            break;
          case '}':
            keybd_event(0, (byte)0xDD, _kbUnicode, UIntPtr.Zero);
            break;
          case '-':
            keybd_event(0, (byte)0xBD, _kbUnicode, UIntPtr.Zero);
            break;
          case '=':
            keybd_event(0, (byte)0xBB, _kbUnicode, UIntPtr.Zero);
            break;
          case '+':
            keybd_event(0, (byte)0xBB, _kbUnicode, UIntPtr.Zero);
            break;
          default:
            keybd_event(0, (byte)key, _kbUnicode, UIntPtr.Zero);
            break;
        }
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }


    /// <summary>
    /// Release a char simulating keyboard input
    /// </summary>
    /// <param name="key">The char to be released</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_ReleaseChar(char key) {
      try {
        //check if the key is a special key
        Thread.Sleep(50);
        switch (key) {
          case ',':
            keybd_event(0, (byte)0xBC, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '.':
            keybd_event(0, (byte)0xBE, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '/':
            keybd_event(0, (byte)0xBF, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '\\':
            keybd_event(0, (byte)0xDC, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case ';':
            keybd_event(0, (byte)0xBA, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '\'':
            keybd_event(0, (byte)0xDE, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '[':
            keybd_event(0, (byte)0xDB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case ']':
            keybd_event(0, (byte)0xDD, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '{':
            keybd_event(0, (byte)0xDB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '}':
            keybd_event(0, (byte)0xDD, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '-':
            keybd_event(0, (byte)0xBD, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '=':
            keybd_event(0, (byte)0xBB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          case '+':
            keybd_event(0, (byte)0xBB, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
          default:
            keybd_event(0, (byte)key, _kbUnicode | 0x0002, UIntPtr.Zero);
            break;
        }
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Type string simulating keyboard input
    /// </summary>
    /// <param name="text">string to type</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_TypeString(String text) { 
      try {
        foreach (char c in text) {
          KB_TypeChar(c);
          
        }
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press a combination of keys together as hotkeys
    /// </summary>
    /// <param name="keys">String of keys to be pressed</param>
    /// <param name="time">The time to hold the keys</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressHotKeys(String[] keys, int time) {
      try {
        foreach (String key in keys) {
          if (key.Length == 1) {
            KB_HoldChar(key[0]);
          }
          else if (key.Length == 2 && key[0] == 'F') {
            KB_HoldFunctionKey(int.Parse(key[1].ToString()));
          }
          else {
            KB_HoldSpecialKey(key);
          }
        }
        Thread.Sleep(time);
        foreach (String key in keys) {
          if (key.Length == 1) {
            KB_ReleaseChar(key[0]);
          }
          else if (key.Length == 2 && key[0] == 'F') {
            KB_ReleaseFunctionKey(int.Parse(key[1].ToString()));
          }
          else {
            KB_ReleaseSpecialKey(key);
          }
        }
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press and hold a special key to help PressHotKeys
    /// </summary>
    /// <param name="key">String of special key to be pressed</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_HoldSpecialKey(String key) {
      Thread.Sleep(50);
      switch (key) {
        case "LShift":
          keybd_event(_kbLShift, 0, 0, UIntPtr.Zero);
          break;
        case "RShift":
          keybd_event(_kbRShift, 0, 0, UIntPtr.Zero);
          break;
        case "LCtrl":
          keybd_event(_kbLCTRL, 0, 0, UIntPtr.Zero);
          break;
        case "RCtrl":
          keybd_event(_kbRCTRL, 0, 0, UIntPtr.Zero);
          break;
        case "LAlt":
          keybd_event(_kbLALT, 0, 0, UIntPtr.Zero);
          break;
        case "RAlt":
          keybd_event(_kbRALT, 0, 0, UIntPtr.Zero);
          break;
        case "Left":
          keybd_event(_kbLeft, 0, 0, UIntPtr.Zero);
          break;
        case "Up":
          keybd_event(_kbUp, 0, 0, UIntPtr.Zero);
          break;
        case "Right":
          keybd_event(_kbRight, 0, 0, UIntPtr.Zero);
          break;
        case "Down":
          keybd_event(_kbDown, 0, 0, UIntPtr.Zero);
          break;
        case "Delete":
          keybd_event(_kbDelete, 0, 0, UIntPtr.Zero);
          break;
        case "Tab":
          keybd_event(_kbTab, 0, 0, UIntPtr.Zero);
          break;
        case "CapsLock":
          keybd_event(_kbCapsLock, 0, 0, UIntPtr.Zero);
          break;
        case "Escape":
          keybd_event(_kbEscape, 0, 0, UIntPtr.Zero);
          break;
        case "Backspace":
          keybd_event(_kbBackspace, 0, 0, UIntPtr.Zero);
          break;
        case "Enter":
          keybd_event(_kbEnter, 0, 0, UIntPtr.Zero);
          break;
        case "Win":
          keybd_event(_kbWin, 0, 0, UIntPtr.Zero);
          break;
        default:
          return -1;
        }
        return 0;
    }

    /// <summary>
    /// Release a special key to help PressHotKeys
    /// </summary>
    /// <param name="key">String of special key to be released</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_ReleaseSpecialKey(String key) {
      Thread.Sleep(50);
      switch (key) {
        case "LShift":
          keybd_event(_kbLShift, 0, 0x0002, UIntPtr.Zero);
          break;
        case "RShift":
          keybd_event(_kbRShift, 0, 0x0002, UIntPtr.Zero);
          break;
        case "LCtrl":
          keybd_event(_kbLCTRL, 0, 0x0002, UIntPtr.Zero);
          break;
        case "RCtrl":
          keybd_event(_kbRCTRL, 0, 0x0002, UIntPtr.Zero);
          break;
        case "LAlt":
          keybd_event(_kbLALT, 0, 0x0002, UIntPtr.Zero);
          break;
        case "RAlt":
          keybd_event(_kbRALT, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Left":
          keybd_event(_kbLeft, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Up":
          keybd_event(_kbUp, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Right":
          keybd_event(_kbRight, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Down":
          keybd_event(_kbDown, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Delete":
          keybd_event(_kbDelete, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Tab":
          keybd_event(_kbTab, 0, 0x0002, UIntPtr.Zero);
          break;
        case "CapsLock":
          keybd_event(_kbCapsLock, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Escape":
          keybd_event(_kbEscape, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Backspace":
          keybd_event(_kbBackspace, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Enter":
          keybd_event(_kbEnter, 0, 0x0002, UIntPtr.Zero);
          break;
        case "Win":
          keybd_event(_kbWin, 0, 0x0002, UIntPtr.Zero);
          break;
        default:
          return -1;
      }
      return 0;
    }

    /// <summary>
    /// Press function key
    /// </summary>
    /// <param name="key">the number of function key to be pressed </param>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressFunctionKey(int key, int time) {
      try {
        Thread.Sleep(50);
        byte functionKey = (byte)(0x70 + key);
        keybd_event(functionKey, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event(functionKey, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Hold function key
    /// </summary>
    /// <param name="key">the number of function key to be pressed </param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_HoldFunctionKey(int key) {
      try {
        Thread.Sleep(50);
        byte functionKey = (byte)(0x70 + key);
        keybd_event(functionKey, 0, 0, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Release function key
    /// </summary>
    /// <param name="key">the number of function key to be pressed </param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_ReleaseFunctionKey(int key) {
      try {
        Thread.Sleep(50);
        byte functionKey = (byte)(0x70 + key);
        keybd_event(functionKey, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press windows key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressWin(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbWin, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbWin, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press escape key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressEscape(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbEscape, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbEscape, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press backspace key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressBackspace(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbBackspace, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbBackspace, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press space key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressSpace(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbSpace, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbSpace, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press enter key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressEnter(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbEnter, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbEnter, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press left shift key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressLShift(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbLShift, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbLShift, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press right shift key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressRShift(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbRShift, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbRShift, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press left control key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressLCTRL(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbLCTRL, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbLCTRL, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press right control key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressRCTRL(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbRCTRL, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbRCTRL, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press left alt key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressLALT(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbLALT, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbLALT, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press right alt key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressRALT(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbRALT, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbRALT, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press left arrow key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressLeft(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbLeft, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbLeft, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press up arrow key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressUp(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbUp, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbUp, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press right arrow key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressRight(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbRight, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbRight, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press down arrow key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressDown(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbDown, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbDown, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press delete key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressDelete(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbDelete, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbDelete, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press tab key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressTab(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbTab, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbTab, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e){
        Console.WriteLine(e.Message);
        return -1;
      }
    }

    /// <summary>
    /// Press caps lock key
    /// </summary>
    /// <param name="time">The time to hold the key</param>
    /// <returns>0 if success, -1 if error</returns>
    public int KB_PressCapsLock(int time) {
      try {
        Thread.Sleep(50);
        keybd_event( _kbCapsLock, 0, 0, UIntPtr.Zero);
        Thread.Sleep(time);
        keybd_event( _kbCapsLock, 0, 0x0002, UIntPtr.Zero);
        return 0;
      }
      catch (Exception e) {
        Console.WriteLine(e.Message);
        return -1;
      }
    }
  }
}