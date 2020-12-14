using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyCsharpExam
{
    class ProgressServing
    {
        ProgressBar bar;
        Label progressLabel;
        
        public ProgressServing(ProgressBar bar, Label progressLabel)
        {
            this.bar = bar;
            this.progressLabel = progressLabel;
        }
    }
}