﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;

namespace WorshipHelperVSTO
{
    class WindowManager
    {
        Application app = Globals.ThisAddIn.Application;

        public DocumentWindow GetPresenterView()
        {
            if (app.Presentations.Count == 0)
            {
                return null;
            }
            
            foreach (DocumentWindow win in app.ActivePresentation.Windows)
            {
                // There is probably a better way...
                if (win.Caption.Contains("Presenter View"))
                {
                    return win;
                }
            }
            return null;
        }
    }
}