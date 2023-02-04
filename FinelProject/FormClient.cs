using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using IronXL;

namespace FinelProject
{
    public partial class FormClient : Form
    {

        string txtFilename;
        bool eastFlg = false;
        bool ethFlg = false;
        string user; // to - Hello , "username"!

        public FormClient(string user)
        {
            this.user = user;
            InitializeComponent();

        }

        private void FormClient_Load(object sender, EventArgs e)
        {
            txthello.Text = "Hello, " + user.ToString() + " !";

            if (radioButton1.Checked)
                radioButton2.Checked = false;
            else
                radioButton1.Checked = false;
        }

        private void logOut_Click(object sender, EventArgs e)
        {
            DialogResult logout = MessageBox.Show("Are you Sure?", "Log Out ", MessageBoxButtons.YesNo);
            switch (logout)
            {
                case DialogResult.Yes:
                    new FormLogin().Show();
                    this.Hide();
                    break;
                case DialogResult.No:
                    break;
            }

        }

        private void addExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" }) //open file .xlsx
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read)) ;
                    vSelect.Visible = true;
                }
            }
        }

        private void resultExcel_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.Default;

            WorkBook workbook = WorkBook.Load(txtFilename);

            WorkSheet sheet = workbook.GetWorkSheet("Blood Test's");

            var sheetResult = workbook.CreateWorkSheet("result's");

            sheetResult["A1"].Value = "First name";
            sheetResult["B1"].Value = "Last name";
            sheetResult["C1"].Value = "ID";
            sheetResult["D1"].Value = "Age";
            sheetResult["E1"].Value = "Smoke?";
            sheetResult["F1"].Value = "Gander";
            sheetResult["G1"].Value = "WBC";
            sheetResult["H1"].Value = "Neut";
            sheetResult["I1"].Value = "Lymph";
            sheetResult["J1"].Value = "RBC";
            sheetResult["K1"].Value = "HCT";
            sheetResult["L1"].Value = "Urea";
            sheetResult["M1"].Value = "Hb";
            sheetResult["N1"].Value = "Crtn";
            sheetResult["O1"].Value = "Iron";
            sheetResult["P1"].Value = "HDL";
            sheetResult["Q1"].Value = "AP";
            sheetResult["R1"].Value = "Diagnosis";
            sheetResult["S1"].Value = "Recommendation";



            //set value to multiple cells
            sheetResult["A2"].Value = txtFname.Text;
            sheetResult["B2"].Value = txtLname.Text;
            sheetResult["C2"].Value = txtID.Text;
            sheetResult["D2"].Value = txtAge.Text;

            if (checkBoxSmoke.Checked)
                sheetResult["E2"].Value = "Yes";
            else
                sheetResult["E2"].Value = "No";

            if (radioButton1.Checked)
                sheetResult["F2"].Value = "Male";
            else
                sheetResult["F2"].Value = "Female";


            sheetResult["G2"].Value = sheet["A2"].Value;
            sheetResult["H2"].Value = sheet["B2"].Value;
            sheetResult["I2"].Value = sheet["C2"].Value;
            sheetResult["J2"].Value = sheet["D2"].Value;
            sheetResult["K2"].Value = sheet["E2"].Value;
            sheetResult["L2"].Value = sheet["F2"].Value;
            sheetResult["M2"].Value = sheet["G2"].Value;
            sheetResult["N2"].Value = sheet["H2"].Value;
            sheetResult["O2"].Value = sheet["I2"].Value;
            sheetResult["P2"].Value = sheet["J2"].Value;
            sheetResult["Q2"].Value = sheet["K2"].Value;
            sheetResult["R2"].Value = "";
            sheetResult["S2"].Value = "";





            int age = Convert.ToInt32(sheetResult["D2"].Value);
            string gender = sheetResult["F2"].Value.ToString();
            int wbc = Convert.ToInt32(sheetResult["G2"].Value);
            int neut = Convert.ToInt32(sheetResult["H2"].Value);
            int lymph = Convert.ToInt32(sheetResult["I2"].Value);
            int rbc = Convert.ToInt32(sheetResult["J2"].Value);
            int hct = Convert.ToInt32(sheetResult["K2"].Value);
            int urea = Convert.ToInt32(sheetResult["L2"].Value);
            int hb = Convert.ToInt32(sheetResult["M2"].Value);
            int crtn = Convert.ToInt32(sheetResult["N2"].Value);
            int iron = Convert.ToInt32(sheetResult["O2"].Value);
            int hdl = Convert.ToInt32(sheetResult["P2"].Value);
            int ap = Convert.ToInt32(sheetResult["Q2"].Value);



            DialogResult df = MessageBox.Show("האם אתה יוצא עדות המזרח?", "שאלה ", MessageBoxButtons.YesNo);
            switch (df)
            {
                case DialogResult.Yes:
                    eastFlg = true;
                    break;
                case DialogResult.No:
                    eastFlg = false;
                    break;
            }

            if (eastFlg == false)
            {
                DialogResult dr = MessageBox.Show("האם אתה יוצא עדות אפריקה?", "שאלה ", MessageBoxButtons.YesNo);
                switch (dr)
                {
                    case DialogResult.Yes:
                        ethFlg = true;
                        break;
                    case DialogResult.No:
                        ethFlg = false;
                        break;
                }

                if (ethFlg == true)
                {
                    DialogResult d = MessageBox.Show("האם אתה יוצא אתיופיה?", "שאלה ", MessageBoxButtons.YesNo);
                    switch (d)
                    {
                        case DialogResult.Yes:
                            ethFlg = true;
                            break;
                        case DialogResult.No:
                            ethFlg = false;
                            break;
                    }

                }
            }

            //WBC
            if (checkWBC(age, wbc)=="high")
            {
                sheetResult["R2"].Value += "אם קיים חום - לרוב זיהום, במקרים נדירים - מחלות דם/סרטן \n";
                sheetResult["S2"].Value += "זיהום - אנטיבוטיקה יעודית/ דימום - להתפנות דחוף לביהח/ סרטן - אנטרקטיניב \n";
            }
            else if(checkWBC(age, wbc) == "low")
            {
                sheetResult["R2"].Value += "מחלה ויראלית, במקרים נדירים - סרטן \n";
                sheetResult["S2"].Value += "מחלה ויראלית - לנוח ביית/ סרטן - אנטרקטיניב \n";
            }

            //Neut

            if(checkNeut(neut)=="high")
            {
                sheetResult["R2"].Value += "מעיד לרוב על זיהום חיידקי \n";
                sheetResult["S2"].Value += "זיהום - אנטיבוטיקה יעודית \n";
            }
            else if(checkNeut(neut)=="low")
            {
                sheetResult["R2"].Value += "הפרעה ביצירת דם - נטייה לזיהום חידקי, במקרים נדירים - סרטן \n";
                sheetResult["S2"].Value += "הפרעה ביצירת דם - כדור 10 מ ג" + " B12 " + "למשך חודש" + "\n";
                sheetResult["S2"].Value += " כדור 5 מ ג של חומצה פולית ביום למשך חודש" + "\n";
                sheetResult["S2"].Value += "זיהום - אנטיבוטיקה יעודית/ סרטן - אנטרקטיניב \n";
            }

            //Lymph

            if (checkLymph(lymph) == "high")
            {
                sheetResult["R2"].Value += "זיהום חיידקי ממושך או על סרטן הלימפומה \n";
                sheetResult["S2"].Value += "זיהום - אנטיבוטיקה יעודית/ סרטן - אנטרקטיניב \n";
            }
            else if (checkLymph(lymph) == "low")
            {
                sheetResult["R2"].Value += " בעיה ביצירת תאי הדם \n";
                sheetResult["S2"].Value += "הפרעה ביצירת דם - כדור 10 מ ג" + " B12 " + "למשך חודש" + "\n";
                sheetResult["S2"].Value += " כדור 5 מ ג של חומצה פולית ביום למשך חודש" + "\n";
            }

            //RBC

            if (checkRBC(rbc) == "high")
            {
                if (checkBoxSmoke.Checked)
                {
                    sheetResult["R2"].Value += "כתוצאה מהעישון \n";
                    sheetResult["S2"].Value += "להפסיק לעשן \n";
                }
                else
                {
                    sheetResult["R2"].Value += "הפרעה במערכת יצור הדם  \n";
                    sheetResult["S2"].Value += "הפרעה ביצירת דם - כדור 10 מ ג" + " B12 " + "למשך חודש" + "\n";
                    sheetResult["S2"].Value += " כדור 5 מ ג של חומצה פולית ביום למשך חודש" + "\n";
                }
            }
            else if (checkRBC(rbc) == "low")
            {
                sheetResult["R2"].Value += "אנמיה / דימומים קשים \n";
                sheetResult["S2"].Value += "אנמיה - שני כדורי 10 מ ג " + " B12 " + " ביום למשך חודש \n";
                sheetResult["S2"].Value += " דימומים קשים - להתפנות בדחיפות לביה ח \n";
            }

            //HCT

            if (checkHCT(hct, gender) == "high")
            {
                if (checkBoxSmoke.Checked)
                {
                    sheetResult["R2"].Value += "כתוצאה מהעישון \n";
                    sheetResult["S2"].Value += "להפסיק לעשן \n";
                }
            }
            else if (checkHCT(hct, gender) == "low")
            {
                sheetResult["R2"].Value += "אנמיה / דימום \n";
                sheetResult["S2"].Value += "אנמיה - שני כדורי 10 מ ג " + " B12 " + " ביום למשך חודש \n";
                sheetResult["S2"].Value += " דימום - להתפנות בדחיפות לביה ח \n";
            }

            //Urea

            if (checkUrea(eastFlg, urea) == "low")
            {
                sheetResult["R2"].Value += " תת תזונה, דיאטה דלת חלבון או מחלת כבד \n";
                sheetResult["S2"].Value += "להפסיק לעשן \n";

            }
            else if (checkUrea(eastFlg, urea) == "high")
            {
                sheetResult["R2"].Value += "מחלות כליה, התייבשות או דיאטה עתירת חלבונים \n";
                sheetResult["S2"].Value += "מחלות כליה -איזון רמת הסוכר בדם , התייבשות - מנוחה מוחלטת בשכיבה, החזרת נוזלים בשתייה \n";
                sheetResult["S2"].Value += "דיאטה - לתאם פגישה עם תזונאי \n";

            }

            //Hb
            if (checkHb(gender, hb, age) == "low")
            {

                sheetResult["R2"].Value += "אנמיה - זו יכולה לנבוע מהפרעה המטולוגית, ממחסור בברזל ומדימומים \n";
                sheetResult["S2"].Value += "הפרעה הומטולוגית - זריקה של הורמון לעידוד ייצור תאי דם אדומים \n";
                sheetResult["S2"].Value += "מחסור ברזל - שני כדורי 10 מ ג " + " B12 " + " ביום למשך חודש \n";
                sheetResult["S2"].Value += "דימומים - להתפנות בדחיפות לביה ח  \n";
            }


            //Crtn
            if (checkCrtn(age, crtn) == "high")
            {
                sheetResult["R2"].Value += "בעיה כלייתית , במקרים נדירים אי ספיקת כליות. שלשולים והקאות/מחלות שריר/צריכה מוגברת של בשר \n";
                sheetResult["S2"].Value += "בעיה כלייתית - איזון רמת הסוכר בדם \n";
                sheetResult["S2"].Value += "מחלות שריר - שני כדורי 5 מ ג של כורכום  " + " C3 " + " של אלטמן ביום למשך חודש \n";
                sheetResult["S2"].Value += "צריכה מוגברת של בשר - לתאם פגישה עם תזונאי \n";
            }
            else if (checkCrtn(age, crtn) == "low")
            {
                sheetResult["R2"].Value += "מסת שריר ירודה/ תת תזונה - לא נצרך מספיק חלבון \n";
                sheetResult["S2"].Value += "לתאם פגישה עם תזונאי \n";
            }

            //Iron

            if (checkIron(iron,gender)=="high")
                {
                    sheetResult["R2"].Value += "עלול להצביע על הרעלת ברזל \n";
                    sheetResult["S2"].Value += "הרעלת ברזל - להתפנות לביה ח  \n";
                }
            else if (checkIron(iron, gender) == "low")
                    {
                    sheetResult["R2"].Value += "תזונה לא מספקת או על עלייה בצורך בברזל / איבוד דם בעקבות דימום \n";
                    sheetResult["S2"].Value += "לתאם פגישה עם תזונאי / דימום - להתפנות בדחיפות לביה ח  \n";

                }

            //HDL
            
            if(checkHDL(ethFlg,hdl,gender)=="low")
            {
                sheetResult["R2"].Value += "סיכון למחלות לב, על היפרליפידמיה (יתר שומנים בדם) או על סוכרת מבוגרים \n";
                sheetResult["S2"].Value += "סיכון למחלת לב - לתאם פגישה עם תזונאי / היפרליפדמיה - לתאם פגישה עם תזונאי, כדור 5 מג של סימוביל ביום למשך שבוע / סכרת מבוגרים - התאמת אינסולין למטופל \n ";
            }


            //Ap

            if (checkAp(eastFlg, ap) == "low")
            {
                sheetResult["R2"].Value += "תזונה לקויה חסרת חלבונים / חוסר בויטמינים \n";
                sheetResult["S2"].Value += "תזונה לקויה - לתאם פגישה עם תזונאי / חוסר בויטמינים - הפניה לבדיקת דם לזיהוי הויטמינים החסרים \n ";
            }
            else if (checkAp(eastFlg, ap) == "high")
            {
                sheetResult["R2"].Value += "מחלות כבד, מחלות בדרכי המרה, הריון, פעילות יתר של בלוטת התריס או שימוש בתרופות שונות \n";
                sheetResult["S2"].Value += "מחלת כבד - הפניה לאבחנה ספציפית לקביעת טיפול \n ";
                sheetResult["S2"].Value += "פעילות יתר בלוטות תריס - " + " Propylthiouracil" + " להקטנת פעילות בלוטות התריס \n";
                sheetResult["S2"].Value += "שימוש בתרופות שונות - הפנייה לרופא המשפחה לצורך בדיקת התאמה בין התרופות \n ";

            }






            Console.WriteLine("The result's are created.");
            createV.Visible = true;

            workbook.Save();

        }

        public string checkWBC(int age, int wbc)
        {
            if (age > 18)
            {
                if (wbc > 11000)
                    return "high";
                else if (wbc < 4500)
                    return "low";
            }
            else if (age >= 4 && age <= 17)
            {
                if (wbc > 15500)
                    return "high";
                else if (wbc < 5500)
                    return "low";
            }
            else if (age >= 0 && age <= 3)
            {
                if (wbc > 17500)
                    return "high";
                else if (wbc < 6000)
                    return "low";
            }
            return "normal";
        }
        public string checkNeut(int neut)
        {
            if (neut < 28)
                return "low";
            else if (neut > 54)
                return "high";
            else
                return "normal";
        }

        public string checkLymph(int lymph)
        {
            if (lymph < 36)
                return "low";
            else if (lymph > 52)
                return "high";
            else
                return "normal";
        }

        public string checkRBC(int rbc)
        {
            if (rbc < 4.5)
                return "low";
            else if (rbc > 6)
                return "high";
            else
                return "normal";
        }

        public string checkHCT(int hct,string gender)
        {
            if (gender == "Male")
            {
                if (hct > 54)
                    return "high";
                else if (hct < 37)
                    return "low";
                else
                    return "normal";
            }
            else
            {
                if (hct > 47)
                    return "high";
                else if (hct < 33)
                    return "low";
                else
                    return "normal";
            }
        }

        public string checkUrea(bool flg, int urea)
        {
            if (flg == false)
            {
                if (urea < 18.7)
                    return "low";
                else if (urea > 47.3)
                    return "high";
                else
                    return "normal";
            }
            else
            {
                if (urea < 17)
                    return "low";
                else if (urea > 43)
                    return "high";
                else
                    return "normal";
            }
        }

        public string checkHb(string gender, int hb, int age)
        {
            if (gender == "Male")
            {
                if (hb < 12)
                    return "low";
            }
            else if (gender == "Female")
            {
                if (hb < 12)
                    return "low";
            }
            else if (age <= 17)
            {
                if (hb < 11.5)
                    return "low";
            }
            return "normal";
        }

        public string checkCrtn(int age, int crtn)
        {
            if (age <= 2)
            {
                if (crtn < 0.2)
                    return "low";
                else if (crtn > 0.5)
                    return "high";

            }
            else if (age <= 17 && age >= 3)
            {
                if (crtn < 0.5)
                    return "low";
                else if (crtn > 1)
                    return "high";

            }
            else if (age <= 59 && age >= 18)
            {
                if (crtn < 0.6)
                    return "low";
                else if (crtn > 1)
                    return "high";
                    
            }
            else if (age <= 60)
            {
                if (crtn < 0.6)
                    return "low";
                else if (crtn > 1.2)
                    return "high";

            }
            return "normal";
        }

        public string checkIron(int iron, string gender)
        {
            if (gender == "Female")
            {
                if (iron > 128)// Female less 20%
                    return "high";
                else if (iron < 48)// Female less 20%
                    return "low";
                else
                    return "normal";
            }
            else
            {
                if (iron > 160)
                    return "high";
                else if (iron < 60)
                    return "low";
                else
                    return "normal";
            }
        }

        public string checkHDL(bool flg, int hdl, string gender)
        {
            if (flg == true)
            {
                if (gender == "Female") //+20%
                {
                    if (hdl < 28.8)
                        return "low";
                }

                else
                {
                    if (hdl < 40.8) //+20%
                        return "low";
                }
                return "normal";
            }
            else
            {
                if (gender == "Female")
                {
                    if (hdl < 24)
                        return "low";
                }

                else
                {
                    if (hdl < 34)
                        return "low";
                }
                return "normal";
            }
        }

        public string checkAp(bool flg, int ap)
        {
            if (flg == true)
            {
                if (ap < 60)
                    return "low";
                else if (ap > 120)
                    return "high";
                else
                    return "normal";
            }
            else
            {
                if (ap < 30)
                    return "low";
                else if (ap > 90)
                    return "high";
                else
                    return "normal";
            }
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
    }
}
