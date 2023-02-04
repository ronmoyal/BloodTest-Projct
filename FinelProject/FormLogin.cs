using IronXL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FinelProject
{
    public partial class FormLogin : Form
    {
        bool userNameFlg = false;
        int userNameIndx;
        string user;

        public FormLogin()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WorkBook workBook = WorkBook.Load(@"C:\Users\IMOE001\source\repos\FinelProject\ExcelFile.xlsx");
            WorkSheet sheet = workBook.GetWorkSheet("Users");
            foreach(var cell in sheet["A2:A100"])
            {
                if (txtUserName.Text == cell.Text)
                {
                    userNameFlg = true;
                    userNameIndx = cell.RowIndex+1;
                    break;
                }
            }
            
            if (userNameFlg==true)
            {
                if(sheet["B"+userNameIndx].ToString()==txtPassword.Text)
                {
                    user = sheet["A" + userNameIndx].ToString();
                    new FormClient(user).Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("The User name or password you entered is incorrect,try again!");
                    txtUserName.Clear();
                    txtPassword.Clear();
                    txtUserName.Focus();
                }
            }

            else
            {
                MessageBox.Show("The User name or password you entered is incorrect,try again!");
                txtUserName.Clear();
                txtPassword.Clear();
                txtUserName.Focus();
            }
            userNameFlg = false;

        }

        private void label2_Click(object sender, EventArgs e)
        {
            DialogResult exit = MessageBox.Show("Are you Sure?", "Exit ", MessageBoxButtons.YesNo);
            switch (exit)
            {
                case DialogResult.Yes:
                    Application.Exit();
                    break;
                case DialogResult.No:
                    break;
            }

            }

            private void label4_Click(object sender, EventArgs e)
        {
            new FormSignup().Show();
                this.Hide();
        }

        private void FormLogin_Load(object sender, EventArgs e)
        {
            txtPassword.PasswordChar = '\u25CF';
        }


    }
}
