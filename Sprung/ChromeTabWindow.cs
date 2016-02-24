using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sprung {
    class ChromeTabWindow : Window {
        public ChromeTabWindow(IntPtr handle, int currentTabIndex, int tabIndex, string title) : base(handle)
        {
            //this.tabIndex = tabIndex;
            //this.tabTitle = title;
            //this.title = title;
            //this.currentTabIndex = currentTabIndex;
        }
    }
}
