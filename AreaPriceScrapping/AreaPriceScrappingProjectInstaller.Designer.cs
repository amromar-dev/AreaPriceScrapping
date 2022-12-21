namespace AreaPrice.Scrapping
{
    partial class AreaPriceScrappingProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.AreaPriceScrappingServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.AreaPriceScrappingServiceInstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // AreaPriceScrappingServiceProcessInstaller
            // 
            this.AreaPriceScrappingServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.AreaPriceScrappingServiceProcessInstaller.Password = null;
            this.AreaPriceScrappingServiceProcessInstaller.Username = null;
            // 
            // AreaPriceScrappingServiceInstaller
            // 
            this.AreaPriceScrappingServiceInstaller.Description = "Service to scrapping some of data from a web site which occurs hourly.";
            this.AreaPriceScrappingServiceInstaller.DisplayName = "Area Price Scrapping ";
            this.AreaPriceScrappingServiceInstaller.ServiceName = "AreaPriceScrapping";
            // 
            // AreaPriceScrappingProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.AreaPriceScrappingServiceProcessInstaller,
            this.AreaPriceScrappingServiceInstaller});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller AreaPriceScrappingServiceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller AreaPriceScrappingServiceInstaller;
    }
}