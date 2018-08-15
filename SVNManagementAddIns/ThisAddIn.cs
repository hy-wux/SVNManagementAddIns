using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SVNManagementAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //hook = new KeyboardHook();
            //hook.InitHook();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //hook.UnHook();
        }

        #region 全局快捷键

        KeyboardHook hook;

        //钩子具体处理逻辑是：
        class KeyboardHook
        {
            #region (invokestuff)
            [DllImport("kernel32.dll")]
            static extern uint GetCurrentThreadId();
            [DllImport("user32.dll")]
            static extern IntPtr SetWindowsHookEx(int code, HookProcKeyboard func, IntPtr hInstance, uint threadID);
            [DllImport("user32.dll")]
            static extern bool UnhookWindowsHookEx(IntPtr hhk);
            [DllImport("user32.dll")]
            static extern int CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
            #endregion

            #region constans
            private const int WH_KEYBOARD = 2;
            private const int HC_ACTION = 0;
            #endregion

            delegate int HookProcKeyboard(int code, IntPtr wParam, IntPtr lParam);
            private HookProcKeyboard KeyboardProcDelegate = null;
            private IntPtr khook;
            bool doing = false;

            public void InitHook()
            {
                uint id = GetCurrentThreadId();
                //init the keyboard hook with the thread id of the Visual Studio IDE   
                this.KeyboardProcDelegate = new HookProcKeyboard(this.KeyboardProc);
                khook = SetWindowsHookEx(WH_KEYBOARD, this.KeyboardProcDelegate, IntPtr.Zero, id);
            }

            public void UnHook()
            {
                if (khook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(khook);
                }
            }

            private int KeyboardProc(int code, IntPtr wParam, IntPtr lParam)
            {
                try
                {
                    if (code != HC_ACTION)
                    {
                        return CallNextHookEx(khook, code, wParam, lParam);
                    }

                    if ((int)wParam == (int)Keys.Z && ((int)lParam & (int)Keys.Alt) != 0)
                    {
                        if (!doing)
                        {
                            doing = true;
                            MessageBox.Show("Captured");
                            doing = false;
                        }
                    }
                }
                catch
                {
                }

                return CallNextHookEx(khook, code, wParam, lParam);
            }
        }

        #endregion

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
