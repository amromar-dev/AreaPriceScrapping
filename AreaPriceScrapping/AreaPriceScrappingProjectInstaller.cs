using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
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
            var exportFolder = Context.Parameters["ExportFolder"];
            var intervalMinutes = Context.Parameters["IntervalMinutes"];

            Context.Parameters["assemblypath"] += $"\" /{exportFolder} /{intervalMinutes}";
            base.OnBeforeInstall(savedState);
        }
    }
}
