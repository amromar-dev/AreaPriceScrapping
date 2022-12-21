using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Configuration.Install;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace AreaPrice.Scrapping
{
    [RunInstaller(true)]
    public partial class AreaPriceScrappingProjectInstaller : Installer
    {
        public AreaPriceScrappingProjectInstaller()
        {
            InitializeComponent();
        }

        protected override void OnBeforeInstall(IDictionary savedState)
        {
            base.OnBeforeInstall(savedState);
        }

        protected override void OnAfterInstall(IDictionary savedState)
        {
            base.OnAfterInstall(savedState);
        }
    }
}
