using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private Control[] controlsArray = new Control[0];
        private Control[] controlsArray1 = new Control[0];

        private async void textBox1_TextChanged(object sender, EventArgs e)
        {
            label3.Visible = true;
            label12.Visible = true;
            button1.Enabled = true;

            int value = 0;
            try
            {
                value = Convert.ToInt32(textBox1.Text);

            }
            catch (System.FormatException)
            {
                string message = "Можна вводить только числовое значение";
                string caption = "Неверное значение";
                string cap = "Дозволено вводити тільки числове значеня";
                string mes = "Невірне значення";
                if (language == true)
                {
                    MessageBox.Show(cap, mes);
                }
                MessageBox.Show(message, caption);
                MessageBox.Show(message, caption);

                textBox1.Text = " ";
                if (textBox1.Enabled == false)
                {
                    textBox1.Enabled = true;
                }

                await PutTaskDelay();

            }
            int lengthArray = value;
            controlsArray = new Control[lengthArray];
            controlsArray1 = new Control[lengthArray];

            for (int i = 0; i < lengthArray; i++)

            {

                controlsArray[i] = new TextBox() { Name = "textBox" + i.ToString(), Location = new Point(95, 120 + (i * 20)), Text = " ", Visible = true };
                this.Controls.Add(controlsArray[i]);


            }
            //------------------------------------------
            for (int j = 0; j < lengthArray; j++)
            {
                controlsArray1[j] = new TextBox() { Name = "textBox" + j.ToString(), Location = new Point(270, 120 + (j * 20)), Text = " ", Visible = true };
                this.Controls.Add(controlsArray1[j]);
            }

            await PutTaskDelay();
            textBox1.Enabled = false;

        }



        private bool Button1Click = false;
        private void button1_Click(object sender, EventArgs e)

        {

            label2.Visible = false;
            label3.Visible = false;
            textBox1.Visible = false;
            button1.Visible = false;
            label12.Visible = false;
            Button1Click = true;
            panel1.Visible = true;
            this.AcceptButton = button1;

        }



        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        async Task PutTaskDelay()
        {
            await Task.Delay(1000);
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }


        private double Chairs;
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Chairs = Convert.ToDouble(textBox2.Text);
            }
            catch (System.FormatException)
            {
                string message = "Можна вводить только числовое значение";
                string caption = "Неверное значение";
                string cap = "Дозволено вводити тільки числове значеня";
                string mes = "Невірне значення";
                if (language == true)
                {
                    MessageBox.Show(cap, mes);
                }
                MessageBox.Show(message, caption);

                textBox2.Text = " ";


            }
        }

        private int PortionsPerDay;
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                PortionsPerDay = Convert.ToInt32(textBox3.Text);
            }
            catch (System.FormatException)
            {
                string message = "Можна вводить только числовое значение";
                string caption = "Неверное значение";
                string cap = "Дозволено вводити тільки числове значеня";
                string mes = "Невірне значення";
                if (language == true)
                {
                    MessageBox.Show(cap, mes);
                }
                MessageBox.Show(message, caption);
                MessageBox.Show(message, caption);

                textBox3.Text = " ";


            }
        }


        private double FindDishesPortions;
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            int percent = Convert.ToInt32(textBox4.Text);
            double finalpercent = (double)percent / 100;
            FindDishesPortions = PortionsPerDay * finalpercent;
        }




        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == String.Empty && textBox4.Text == String.Empty && textBox12.Text == String.Empty)
            {
                if (language == true)
                {
                    string message1 = "Введіть значения";
                    MessageBox.Show(message1);
                }
                string message = "Введите значения";
                MessageBox.Show(message);
            }
            else {
                groupBox1.Visible = true;
                textBox7.Text = Convert.ToString((int)FindDishesPortions);
                textBox8.Text = Convert.ToString((int)FindSecondDishesPortions);
                textBox9.Text = Convert.ToString((int)FindThirdDishesPortions);
                textBox10.Text = Convert.ToString((int)Dish);
                button5.Visible = true;
            }
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }


        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            string name = textBox11.Text;
            label17.Text = name;

        }

        private double FindSecondDishesPortions;
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            int percent = Convert.ToInt32(textBox5.Text);
            double finalpercent = (double)percent / 100;

            FindSecondDishesPortions = FindDishesPortions * finalpercent;

        }

        private double FindThirdDishesPortions;
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            int percent = Convert.ToInt32(textBox5.Text);
            double finalpercent = (double)percent / 100;

            FindThirdDishesPortions = FindSecondDishesPortions * finalpercent;
        }

        private bool Check1 = false;
        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            textBox5.Visible = true;
            textBox8.Visible = true;
            label19.Visible = true;
            Check1 = true;
            label23.Visible = true;
            textBox13.Visible = true;
            label11.Visible = true;
        }

        private bool Check2 = false;
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            label13.Visible = true;
            textBox6.Visible = true;
            textBox9.Visible = true;
            label20.Visible = true;
            Check2 = true;
            textBox14.Visible = true;
        }


        double Dish;
        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            int menu = Convert.ToInt32(textBox12.Text);
            double find = 0;
            if (Check1 == false && Check2 == false)
            {
                find = FindDishesPortions;
            }
            else if (Check1 == true && Check2 == false)
            {
                find = FindSecondDishesPortions;
            }
            else if (Check1 == true && Check2 == true)
            {
                find = FindThirdDishesPortions;
            }
            Dish = find / menu;
        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            string name = textBox13.Text;
            label18.Text = name;
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            string name = textBox14.Text;
            label24.Text = name;
        }


        private void обнулитьЗначенияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox2.Text = "0";
            textBox3.Text = "0";
            textBox4.Text = "0";
            textBox5.Text = "0";
            textBox6.Text = "0";
            textBox7.Text = "0";
            textBox8.Text = "0";
            textBox9.Text = "0";
            textBox10.Text = "0";
            textBox11.Text = " ";
            textBox12.Text = "0";
            textBox13.Text = " ";
            textBox14.Text = " ";
        }

        private Control[] controlsArrayResult = new Control[0];
        private Control[] controlsArrayResult1 = new Control[0];



        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }




        private void button5_Click(object sender, EventArgs e)
        {



            panel1.Visible = false;

            string[] result = new string[controlsArray1.Length];
            int[] resultt = new int[controlsArray.Length];
            int[] resulttt = new int[controlsArray.Length];
            int[] portion = new int[controlsArray.Length];
            for (int h = 0; h < controlsArray1.Length; h++)
            {
                portion[h] = 4;
            }
            int value = 0;
            value = Convert.ToInt32(textBox1.Text); //Бере значення зі старого текстбокса
            int lengthArray = value;
            controlsArrayResult = new Control[lengthArray];
            controlsArrayResult1 = new Control[lengthArray];


            for (int l = 0; l < controlsArray1.Length; l++)
            {
                result[l] = controlsArray1[l].Text;
            }

            for (int k = 0; k < result.Length; k++)
            {
                resultt[k] = Convert.ToInt32(result[k]);
            }
            for (int x = 0; x < resultt.Length; x++)
            {
                resulttt[x] = resultt[x] / portion[x] * (int)Dish;

            }


            for (int i = 0; i < lengthArray; i++)
            {
                controlsArrayResult[i] = new TextBox() { Name = "textBox" + i.ToString(), Location = new Point(95, 120 + (i * 20)), Text = controlsArray[i].Text, Visible = true };
                this.Controls.Add(controlsArrayResult[i]);
            }

            for (int j = 0; j < lengthArray; j++)
            {
                controlsArrayResult1[j] = new TextBox() { Name = "textBox" + j.ToString(), Location = new Point(370, 120 + (j * 20)), Text = resulttt[j].ToString(), Visible = true };
                this.Controls.Add(controlsArrayResult1[j]);
            }
            label3.Visible = true;
            label12.Visible = true;
            label16.Visible = true;
            label25.Visible = true;
            button5.Visible = false;
            button2.Visible = true;
            button4.Visible = true;
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 fm = new Form2();
            fm.Show();

        }

        public static bool language = false;


        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }



        private void українськаToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            language = true;

            label2.Text = "Введіть кількість інгрідієнтів";
            label3.Text = "Найменування інгрідієнтів\n (Приклад: Картопля)";
            label12.Text = "Маса брутто (г)\n (Приклад: 27, 8)";
            label25.Text = "Маса брутто на 4\n порції по 250г";
            label16.Text = "Результати обчислення округлені до цілих чисел,\n з метою приближення результатів до реальних чисел ";
            label15.Text = "Загальна кількість порцій  ";
            label19.Text = "Кількість порцій  ";
            label20.Text = "Кількість порцій  ";
            label22.Text = "Кількість порцій шуканої страви на одне найменування  ";
            label21.Text = "Введіть кількість шуканих страв в меню  ";
            label4.Text = "Введіть кількість людей на місце  ";
            label5.Text = "Введіть кількість порцій на день  ";
            label14.Text = "Введіть назву шуканих страв\n (Приклад:Заправних)  ";
            checkBox1.Text = "Треба вирахувати % з них";
            checkBox2.Text = "Треба вирахувати % з них";
            label23.Text = "   Назва :";
            label9.Text = "людей";
            groupBox1.Text = "Результати";
            button5.Text = "Показати техкарту";
            параметрыToolStripMenuItem.Text = "Параметри";
            обнулитьЗначенияToolStripMenuItem.Text = "Обнулити значення";
            языкToolStripMenuItem.Text = "Мова";
            оПрограммеToolStripMenuItem.Text = "Про програму";
            выходToolStripMenuItem.Text = "Вихід";
            версия11ToolStripMenuItem.Text = "Версія 1.1";
            button2.Text = "Вихід";
        }

        private void русскийToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            language = false;

            label2.Text = "Введите количество ингридиентов";
            label3.Text = "Найменование ингридиентов\n (Пример: Картошка)";
            label12.Text = "Маса брутто (г)\n (Пример: 27, 8)";
            label25.Text = "Масса брутто на 4\n порции по 250г";
            label16.Text = "Результаты вычислений округлены до целых чисел,\n с целью приближения результатов к реальным числам ";
            label15.Text = "Общее количество порций    ";
            label19.Text = "Количество порций  ";
            label20.Text = "Количество порций  ";
            label22.Text = "Количество  порций искомого блюда на найменование  ";
            label21.Text = "Введите количество искомых блюд в меню  ";
            label4.Text = "Введите количество людей на место   ";
            label5.Text = "Введите количество порций на день  ";
            label14.Text = "Введите название искомых блюд\n (Пример:Заправных)  ";
            checkBox1.Text = "Нужно вычислить % с них";
            checkBox2.Text = "Нужно вычислить % с них";
            label23.Text = "Название :";
            label9.Text = "человек";
            groupBox1.Text = "Результаты";
            button5.Text = "Показать техкарту";
            параметрыToolStripMenuItem.Text = "Параметры";
            обнулитьЗначенияToolStripMenuItem.Text = "Обнулить значения";
            языкToolStripMenuItem.Text = "Язык";
            оПрограммеToolStripMenuItem.Text = "О программе";
            выходToolStripMenuItem.Text = "Выход";
            версия11ToolStripMenuItem.Text = "Версия 1.1";
            button2.Text = "Выход";
        }

        private void button4_Click(object sender, EventArgs e)
        {

            button5.Visible = false;

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                //Manipulate a variable number of columns for Quarterly Sales Data.
                DisplayQuarterlySales(oSheet);

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private Application application;
        private Excel.Workbook workBook;
        private Excel.Worksheet worksheet;

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            Excel._Workbook oWB;
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            String sMsg;
            int iNumQtrs;

            //Determine how many quarters to display data for.
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = String.Concat(sMsg, iNumQtrs);
                sMsg = String.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                    MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = String.Concat(sMsg, iNumQtrs);
            sMsg = String.Concat(sMsg, " quarter(s).");

            MessageBox.Show(sMsg, "Quarterly Sales");

            //Starting at E1, fill headers for the number of columns selected.
            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            //Change the Orientation and WrapText properties for the headers.
            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            //Fill the interior color of the headers.
            oResizeRange.Interior.ColorIndex = 36;

            //Fill the columns with a formula and apply a number format.
            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            //Apply borders to the Sales data and headers.
            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Add a Totals formula for the sales data and apply a border.
            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(E2:E6)";
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
                = Excel.XlLineStyle.xlDouble;
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
                = Excel.XlBorderWeight.xlThick;

            //Add a Chart for the selected data.
            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);

            //Use the ChartWizard to create a new chart from the selected data.
            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
                Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
                Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            //Move the chart so as not to cover your data.
            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;

        }
    }

}
