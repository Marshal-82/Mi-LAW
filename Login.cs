using System;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.Protocols;
using System.Net;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MetroFramework;

namespace WindowsFormsApplication1
{
    public partial class Login : MetroFramework.Forms.MetroForm
    {
        public static string UserN;
        
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            PasswordInput.PasswordChar = '\u25CF';

        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string User = UserInput.Text;
            string Password = PasswordInput.Text;
            if (ValidateCredentials() == true)
            {
                UserN = UserInput.Text.ToString();
                this.DialogResult = DialogResult.OK;
            }

            else
            {
                MetroMessageBox.Show(this, Environment.NewLine + "Usuario o Password no validos, contacta a tu Supervisor", "Error de Acceso", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None;
                Application.Exit();
            }
        }
            
        public bool ValidateCredentials()
        {
            bool validation;
            try
            {
                LdapConnection lcon = new LdapConnection
                (new LdapDirectoryIdentifier((string)null, false, false));
                NetworkCredential nc = new NetworkCredential(UserInput.Text,
                                       PasswordInput.Text, Environment.UserDomainName);
                lcon.Credential = nc;
                lcon.AuthType = AuthType.Negotiate;
                lcon.Bind(nc);
                validation = true;
            }
            catch (LdapException)
            {
                validation = false;
            }
            return validation;
        }
    }
}
