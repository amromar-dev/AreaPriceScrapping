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
    public partial class AreaPriceScrappingProjectInstaller : System.Configuration.Install.Installer
    {
        public AreaPriceScrappingProjectInstaller()
        {
            InitializeComponent();
        }
    }
}
