﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sprung
{
    class FirefoxTabWindow : Window
    {

        protected int tabIndex;
        protected String tabTitle;
        public int currentTabIndex;

        public FirefoxTabWindow(IntPtr handle, int currentTabIndex, int tabIndex, string title) : base(handle)
        {
            this.tabIndex = tabIndex;
            this.tabTitle = title;
            this.title = title;
            this.currentTabIndex = currentTabIndex;
        }

        public int getTabIndex()
        {
            return tabIndex;
        }

        public override void SendToFront()
        {
            base.SendToFront();
            int changeVector = tabIndex - currentTabIndex;
            int tabChanges = Math.Abs(changeVector);
            int direction = Math.Sign(changeVector);
            for (int i = 0; i < tabChanges; i++)
            {
                if (direction < 0)
                {
                    SendKeys.Send("^{PGUP}");
                    Debug.WriteLine("Sent PGUP");
                }
                else
                {
                    SendKeys.Send("^{PGDN}");
                    Debug.WriteLine("Sent PGDN");
                }
            }
        }

    }
}
