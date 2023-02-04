using IronXL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FinelProject
{
    public partial class FormSignup : Form
    {

        bool userOK=false;
        bool passOk=false;
        bool idOk = false;

        public FormSignup()
        {
            InitializeComponent();
        }

        private void FormSignup_Load(object sender, EventArgs e)
        {
            txtPassword2.PasswordChar = '\u25CF';
            txtPassword1.PasswordChar = '\u25CF';
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //userName

            //length
            if (txtUserName.TextLength < 6 || txtUserName.TextLength > 8)
            {
                MessageBox.Show("Username must be 6-8 characters long");
                txtUserName.Clear();
            }
            //digits
            else if (maxTwoDigits(txtUserName.Text) > 2)
            {
                MessageBox.Show("Username must contain maximum 2 numbers");
                txtUserName.Clear();
            }
            //letters and numbers only
            else if (!Regex.IsMatch(txtUserName.Text, "[a-zA-Z]"))
            {
                MessageBox.Show("Username can only contain numbers and letters in English");
                txtUserName.Clear();
            }
            else
                userOK=true;

            //PassWord

            //length
            if (txtPassword2.TextLength < 8 || txtPassword2.TextLength > 10)
            {
                MessageBox.Show("Password must be 8-10 characters long");
                txtPassword2.Clear();
                txtPassword1.Clear();
            }

            //digits
            else if(charDigitSpaciel(txtPassword2.Text)==false)
            {
                MessageBox.Show("Password must contain at least - one letter, one digit, one spaciel character");
                txtPassword2.Clear();
                txtPassword1.Clear();
            }
            
            //confirm
            else if (txtPassword2.Text != txtPassword1.Text)
            {
                MessageBox.Show("Password's are not equals");
                txtPassword1.Clear();
            }
            else
                passOk = true;

            //id 
            if (idNumber.Text.Length != 9)
            {
                MessageBox.Show("Invalid length Id");
                idNumber.Clear();
                return;
            }
            else
                idOk = true;

            if(passOk==true&&userOK==true&&idOk==true)
            {
                WorkBook workBook = WorkBook.Load(@"C:\Users\IMOE001\source\repos\FinelProject\ExcelFile.xlsx");
                WorkSheet sheet = workBook.GetWorkSheet("Users");
                foreach(var cell in sheet["A2:A100"])
                {
                    if(cell.Text ==txtUserName.Text)
                    {
                        MessageBox.Show("This username already exist's");
                        txtUserName.Clear();
                        return;

                    }
                    if (cell.Text=="")
                    {
                        cell.Text = txtUserName.Text;
                        break;
                    }
                }
                foreach (var cell in sheet["B2:B100"])
                {
                    if (cell.Text == "")
                    {
                        cell.Text = txtPassword2.Text;
                        break;
                    }
                }
                foreach (var cell in sheet["C2:C100"])
                {
                    if (cell.Text == "")
                    {
                        cell.Text = idNumber.Text;
                        break;
                    }
                }

                MessageBox.Show("User created successfully!");

                new FormLogin().Show();
                this.Hide();

                workBook.Save();
            }

            else
            {
                userOK = false;
                passOk = false;
                idOk = false;
            }


        }

        private void label4_Click(object sender, EventArgs e)
        {
            new FormLogin().Show();
            this.Hide();
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


        private bool charDigitSpaciel(string passWord)
        {
            bool charFlag = false, digitFlag = false, SpacielFlag = false;

            foreach (char c in passWord)
            {
                if (Regex.IsMatch(passWord, "[a-zA-Z]"))
                    charFlag = true;
                if (c > '0' && c < '9')
                    digitFlag = true;
                if (Regex.IsMatch(passWord, "[~!@#$%^&*()_<>]"))
                    SpacielFlag = true;
            }

            if (charFlag == true && digitFlag == true && SpacielFlag == true)
                return true;
            return false;
        }

        private int maxTwoDigits(string userName)
        {
            int count = 0;
            foreach (char c in userName)
            {
                if (c > '0' && c < '9')
                    count++;
            }
            return count;
        }

    }



}
